"""数据库连接和初始化模块"""
import aiosqlite
import os
from pathlib import Path

# 数据目录
DATA_DIR = Path(__file__).parent.parent / "data"
SYSTEM_NAME = os.getenv("SYSTEM_NAME", "共享表格")
DB_PATH = DATA_DIR / f"{SYSTEM_NAME}.db"
SHEETS_DIR = DATA_DIR / "sheets"
LOGS_DIR = DATA_DIR / "logs"

# 确保目录存在
DATA_DIR.mkdir(exist_ok=True)
SHEETS_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)


async def get_db():
    """获取数据库连接"""
    db = await aiosqlite.connect(DB_PATH)
    db.row_factory = aiosqlite.Row
    return db


async def init_db():
    """初始化数据库表"""
    async with aiosqlite.connect(DB_PATH) as db:
        # 密钥表
        await db.execute("""
            CREATE TABLE IF NOT EXISTS sheet_keys (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                key TEXT UNIQUE NOT NULL,
                name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # 管理员表
        await db.execute("""
            CREATE TABLE IF NOT EXISTS admins (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # 在线用户表（用于追踪协作用户）
        await db.execute("""
            CREATE TABLE IF NOT EXISTS online_users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sheet_key TEXT NOT NULL,
                user_id TEXT NOT NULL,
                ip_address TEXT,
                mac_address TEXT,
                connected_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (sheet_key) REFERENCES sheet_keys(key)
            )
        """)

        await db.commit()

        # 检查是否有默认管理员，没有则创建
        cursor = await db.execute("SELECT COUNT(*) FROM admins")
        count = await cursor.fetchone()
        if count[0] == 0:
            import hashlib
            default_password = hashlib.sha256("admin123".encode()).hexdigest()
            await db.execute(
                "INSERT INTO admins (username, password_hash) VALUES (?, ?)",
                ("admin", default_password)
            )
            await db.commit()
            print("默认管理员已创建: admin / admin123")
