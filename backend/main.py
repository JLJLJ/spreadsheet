"""共享表格 - 多人共享编辑Excel表格 Web应用"""
import os
import uuid
import hashlib
import socket
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv

from fastapi import FastAPI, WebSocket, WebSocketDisconnect, HTTPException, UploadFile, File, Request, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, RedirectResponse, PlainTextResponse
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware

from database import init_db, get_db, SHEETS_DIR
from models import AuthRequest, AuthResponse, SheetKeyCreate
from excel_handler import create_empty_sheet, load_sheet_data, import_excel
from websocket_manager import manager

# 加载.env配置
load_dotenv(Path(__file__).parent.parent / ".env")
ADMIN_KEY = os.getenv("ADMIN_KEY", "admin123456")
SERVER_PORT = os.getenv("SERVER_PORT", "8000")
HISTORY_COUNT = int(os.getenv("HISTORY_COUNT", "20"))
SYNC_INTERVAL = int(os.getenv("SYNC_INTERVAL", "3000"))
IP_WHITELIST = os.getenv("IP_WHITELIST", "21.,127.0.0.1")
SYSTEM_NAME = os.getenv("SYSTEM_NAME", "共享表格")
AUTHOR_INFO = os.getenv("AUTHOR_INFO", "未知作者")


def get_local_ip() -> str:
    """获取本机IPv4地址"""
    try:
        # 创建一个UDP socket连接到外部地址来获取本机IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        # 备用方法
        try:
            hostname = socket.gethostname()
            ip = socket.gethostbyname(hostname)
            return ip
        except Exception:
            return "127.0.0.1"


# 处理SERVER_IP配置
_server_ip_config = os.getenv("SERVER_IP", "NONE")
if _server_ip_config.upper() == "NONE":
    SERVER_IP = get_local_ip()
else:
    SERVER_IP = _server_ip_config


def check_ip_whitelist(client_ip: str) -> bool:
    """检查IP是否在白名单中"""
    if IP_WHITELIST == "*":
        return True

    whitelist = [ip.strip() for ip in IP_WHITELIST.split(",") if ip.strip()]

    for pattern in whitelist:
        if client_ip.startswith(pattern) or client_ip == pattern:
            return True

    return False


class IPWhitelistMiddleware(BaseHTTPMiddleware):
    """IP白名单验证中间件"""

    async def dispatch(self, request: Request, call_next):
        # 获取客户端IP
        client_ip = request.client.host if request.client else "unknown"

        # 检查X-Forwarded-For头（用于代理情况）
        forwarded_for = request.headers.get("X-Forwarded-For")
        if forwarded_for:
            client_ip = forwarded_for.split(",")[0].strip()

        # 静态资源不检查
        if request.url.path.startswith("/static"):
            return await call_next(request)

        # 检查白名单
        if not check_ip_whitelist(client_ip):
            return PlainTextResponse(
                f"Access denied. Your IP ({client_ip}) is not in the whitelist.",
                status_code=403
            )

        return await call_next(request)

app = FastAPI(title=SYSTEM_NAME, description="多人共享编辑Excel表格")

# CORS配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# IP白名单中间件
app.add_middleware(IPWhitelistMiddleware)

# 静态文件
FRONTEND_DIR = Path(__file__).parent.parent / "frontend"
app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")


@app.on_event("startup")
async def startup():
    """启动时初始化数据库"""
    await init_db()
    print("=" * 50)
    print(SYSTEM_NAME)
    print(f"v1.0  {AUTHOR_INFO}")
    print("=" * 50)
    print(f"服务器IP: {SERVER_IP}")
    print(f"IP白名单: {IP_WHITELIST}")
    print(f"前端目录: {FRONTEND_DIR}")
    print(f"数据目录: {SHEETS_DIR}")


# ==================== 页面路由 ====================

@app.get("/", response_class=HTMLResponse)
async def index():
    """根路径 - 返回404（没有主页）"""
    raise HTTPException(status_code=404, detail="页面不存在。请通过管理员提供的链接访问表格。")


@app.get("/sheet/{key}", response_class=HTMLResponse)
async def sheet_page(key: str):
    """通过密钥直接访问表格编辑器"""
    # 验证密钥是否有效
    db = await get_db()
    try:
        cursor = await db.execute(
            "SELECT key FROM sheet_keys WHERE key = ?",
            (key,)
        )
        row = await cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="表格不存在或密钥无效")
        return FileResponse(FRONTEND_DIR / "editor.html")
    finally:
        await db.close()


@app.get("/admin/{admin_key}", response_class=HTMLResponse)
async def admin_page(admin_key: str):
    """通过密钥访问管理员页面"""
    if admin_key != ADMIN_KEY:
        raise HTTPException(status_code=403, detail="管理员密钥无效")
    return FileResponse(FRONTEND_DIR / "admin.html")


# ==================== API路由 ====================

@app.get("/api/config")
async def get_config():
    """获取系统配置信息（包含服务器IP和端口，用于生成访问链接）"""
    return {
        "systemName": SYSTEM_NAME,
        "authorInfo": AUTHOR_INFO,
        "server_ip": SERVER_IP,
        "server_port": SERVER_PORT,
        "history_count": HISTORY_COUNT,
        "sync_interval": SYNC_INTERVAL
    }


@app.get("/api/sheet/{key}/history")
async def get_sheet_history(key: str):
    """获取表格的修改历史"""
    history = manager.get_history(key, HISTORY_COUNT)
    return {"history": history}


@app.post("/api/auth", response_model=AuthResponse)
async def authenticate(auth: AuthRequest):
    """验证密钥 (POST方式)"""
    return await verify_key(auth.key)


@app.get("/api/auth", response_model=AuthResponse)
async def authenticate_get(key: str):
    """验证密钥 (GET方式) - 支持URL参数直接访问"""
    return await verify_key(key)


async def verify_key(key: str) -> AuthResponse:
    """验证密钥的核心逻辑"""
    db = await get_db()
    try:
        cursor = await db.execute(
            "SELECT key, name, file_path FROM sheet_keys WHERE key = ?",
            (key,)
        )
        row = await cursor.fetchone()

        if row:
            return AuthResponse(
                success=True,
                message="验证成功",
                sheet_name=row["name"],
                token=key  # 简化处理，直接用key作为token
            )
        else:
            return AuthResponse(
                success=False,
                message="密钥无效"
            )
    finally:
        await db.close()


@app.get("/api/sheet/{key}")
async def get_sheet(key: str):
    """获取表格数据"""
    db = await get_db()
    try:
        cursor = await db.execute(
            "SELECT file_path FROM sheet_keys WHERE key = ?",
            (key,)
        )
        row = await cursor.fetchone()

        if not row:
            raise HTTPException(status_code=404, detail="表格不存在")

        file_path = row["file_path"]
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="表格文件不存在")

        sheet_data = load_sheet_data(file_path)
        return sheet_data
    finally:
        await db.close()


@app.get("/api/sheet/{key}/export")
async def export_sheet(key: str):
    """导出Excel文件"""
    db = await get_db()
    try:
        cursor = await db.execute(
            "SELECT name, file_path FROM sheet_keys WHERE key = ?",
            (key,)
        )
        row = await cursor.fetchone()

        if not row:
            raise HTTPException(status_code=404, detail="表格不存在")

        file_path = row["file_path"]
        file_name = row["name"]

        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="表格文件不存在")

        return FileResponse(
            file_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=f"{file_name}.xlsx"
        )
    finally:
        await db.close()


# ==================== 管理员API ====================

@app.get("/api/admin/verify")
async def verify_admin_key(key: str):
    """验证管理员密钥"""
    if key == ADMIN_KEY:
        return {"success": True}
    else:
        return {"success": False, "message": "管理员密钥无效"}


@app.get("/api/admin/keys")
async def list_keys():
    """列出所有密钥"""
    db = await get_db()
    try:
        cursor = await db.execute(
            "SELECT key, name, file_path, created_at, updated_at FROM sheet_keys ORDER BY created_at DESC"
        )
        rows = await cursor.fetchall()
        keys = []
        for row in rows:
            online_count = len(manager.get_online_users(row["key"]))
            keys.append({
                "key": row["key"],
                "name": row["name"],
                "file_path": row["file_path"],
                "created_at": row["created_at"],
                "updated_at": row["updated_at"],
                "online_users": online_count
            })
        return {"keys": keys}
    finally:
        await db.close()


@app.post("/api/admin/keys")
async def create_key(
    name: str = Form(...),
    key: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None)
):
    """创建新密钥"""
    db = await get_db()
    try:
        # 生成密钥
        if not key:
            key = str(uuid.uuid4())[:8].upper()

        # 检查密钥是否已存在
        cursor = await db.execute("SELECT key FROM sheet_keys WHERE key = ?", (key,))
        if await cursor.fetchone():
            raise HTTPException(status_code=400, detail="密钥已存在")

        # 创建Excel文件
        if file:
            # 上传的文件
            temp_path = SHEETS_DIR / f"temp_{uuid.uuid4()}.xlsx"
            content = await file.read()
            with open(temp_path, "wb") as f:
                f.write(content)
            file_path = import_excel(str(temp_path), key)
            os.remove(temp_path)
        else:
            # 创建空白表格
            file_path = create_empty_sheet(key)

        # 保存到数据库
        await db.execute(
            "INSERT INTO sheet_keys (key, name, file_path) VALUES (?, ?, ?)",
            (key, name, file_path)
        )
        await db.commit()

        return {"success": True, "key": key, "name": name}
    finally:
        await db.close()


@app.delete("/api/admin/keys/{key}")
async def delete_key(key: str):
    """删除密钥"""
    db = await get_db()
    try:
        # 获取文件路径
        cursor = await db.execute(
            "SELECT file_path FROM sheet_keys WHERE key = ?",
            (key,)
        )
        row = await cursor.fetchone()

        if not row:
            raise HTTPException(status_code=404, detail="密钥不存在")

        # 删除文件
        file_path = row["file_path"]
        if os.path.exists(file_path):
            os.remove(file_path)

        # 删除数据库记录
        await db.execute("DELETE FROM sheet_keys WHERE key = ?", (key,))
        await db.commit()

        return {"success": True}
    finally:
        await db.close()


@app.get("/api/admin/keys/{key}/users")
async def get_key_users(key: str):
    """获取某密钥的在线用户"""
    users = manager.get_online_users(key)
    return {"users": users}


# ==================== WebSocket ====================

@app.websocket("/ws/{key}")
async def websocket_endpoint(websocket: WebSocket, key: str):
    """WebSocket连接端点"""
    # 验证密钥
    db = await get_db()
    try:
        cursor = await db.execute(
            "SELECT file_path FROM sheet_keys WHERE key = ?",
            (key,)
        )
        row = await cursor.fetchone()
        if not row:
            await websocket.close(code=4001, reason="无效的密钥")
            return
        file_path = row["file_path"]
    finally:
        await db.close()

    # 获取客户端信息
    client = websocket.client
    ip_address = client.host if client else "unknown"

    # 从查询参数获取MAC地址（前端发送）
    mac_address = websocket.query_params.get("mac", "unknown")

    # 使用MAC+IP作为唯一用户标识（同一用户重新进入不会被识别为新用户）
    user_id = f"{mac_address}_{ip_address}"

    # 建立连接
    await manager.connect(websocket, key, user_id, ip_address, mac_address, file_path)

    try:
        while True:
            message = await websocket.receive_text()
            await manager.process_message(key, user_id, message)
    except WebSocketDisconnect:
        await manager.notify_disconnect(key, user_id)
        manager.disconnect(key, user_id)


# ==================== 启动入口 ====================

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
