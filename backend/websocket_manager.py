"""WebSocket连接管理器 - 实现多人实时协作"""
import json
import asyncio
import os
from typing import Dict, List, Set, Optional
from fastapi import WebSocket
from datetime import datetime

from excel_handler import update_cell, batch_update_cells, batch_update_dimensions
from database import LOGS_DIR


def get_log_file_path() -> str:
    """获取今天的日志文件路径"""
    today = datetime.now().strftime("%Y-%m-%d")
    return str(LOGS_DIR / f"{today}.log")


def log_user_action(user_id: str, display_name: str, sheet_key: str,
                    action_type: str, details: dict):
    """记录用户操作到日志文件"""
    log_path = get_log_file_path()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]

    log_entry = {
        "timestamp": timestamp,
        "user_id": user_id,
        "display_name": display_name,
        "sheet_key": sheet_key,
        "action": action_type,
        "details": details
    }

    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(log_entry, ensure_ascii=False) + "\n")
    except Exception as e:
        print(f"写入日志失败: {e}")


class ConnectionManager:
    """WebSocket连接管理器"""

    def __init__(self):
        # 按表格密钥分组的活动连接: {sheet_key: {user_id: WebSocket}}
        self.active_connections: Dict[str, Dict[str, WebSocket]] = {}
        # 用户信息: {user_id: {ip, mac, sheet_key, connected_at}}
        self.user_info: Dict[str, Dict] = {}
        # 每个表格的文件路径
        self.sheet_paths: Dict[str, str] = {}
        # 更新队列，用于批量保存
        self.update_queues: Dict[str, List[Dict]] = {}
        # 保存任务
        self.save_tasks: Dict[str, asyncio.Task] = {}
        # 修改历史记录: {sheet_key: [history_entries]}
        self.edit_history: Dict[str, List[Dict]] = {}
        # 历史记录最大条数（默认100，防止内存溢出）
        self.max_history = 100

    async def connect(self, websocket: WebSocket, sheet_key: str, user_id: str,
                      ip_address: str, mac_address: str, file_path: str):
        """建立WebSocket连接"""
        await websocket.accept()

        if sheet_key not in self.active_connections:
            self.active_connections[sheet_key] = {}
            self.update_queues[sheet_key] = []

        # 检查是否是同一用户重新连接
        is_reconnect = user_id in self.active_connections.get(sheet_key, {})

        self.active_connections[sheet_key][user_id] = websocket
        self.sheet_paths[sheet_key] = file_path

        # 只在首次连接时更新用户信息的connected_at
        if user_id not in self.user_info:
            self.user_info[user_id] = {
                "ip": ip_address,
                "mac": mac_address,
                "sheet_key": sheet_key,
                "connected_at": datetime.now().isoformat(),
                "display_name": f"{mac_address}@{ip_address}"
            }
        else:
            # 重新连接时只更新sheet_key
            self.user_info[user_id]["sheet_key"] = sheet_key

        # 只有新用户才通知其他人
        if not is_reconnect:
            await self.broadcast_to_sheet(sheet_key, {
                "type": "user_join",
                "user_id": user_id,
                "display_name": self.user_info[user_id]["display_name"],
                "online_users": self.get_online_users(sheet_key)
            }, exclude=user_id)

        # 发送当前在线用户列表给连接的用户
        await self.send_personal(websocket, {
            "type": "connected",
            "user_id": user_id,
            "display_name": self.user_info[user_id]["display_name"],
            "online_users": self.get_online_users(sheet_key)
        })

    def disconnect(self, sheet_key: str, user_id: str):
        """断开WebSocket连接"""
        if sheet_key in self.active_connections:
            if user_id in self.active_connections[sheet_key]:
                del self.active_connections[sheet_key][user_id]

            # 如果没有用户了，清理资源
            if not self.active_connections[sheet_key]:
                del self.active_connections[sheet_key]
                if sheet_key in self.update_queues:
                    del self.update_queues[sheet_key]

        if user_id in self.user_info:
            del self.user_info[user_id]

    async def notify_disconnect(self, sheet_key: str, user_id: str):
        """通知其他用户有用户离开"""
        display_name = self.user_info.get(user_id, {}).get("display_name", user_id)
        await self.broadcast_to_sheet(sheet_key, {
            "type": "user_leave",
            "user_id": user_id,
            "display_name": display_name,
            "online_users": self.get_online_users(sheet_key)
        }, exclude=user_id)

    def get_online_users(self, sheet_key: str) -> List[Dict]:
        """获取某表格的在线用户列表"""
        users = []
        if sheet_key in self.active_connections:
            for user_id in self.active_connections[sheet_key]:
                info = self.user_info.get(user_id, {})
                users.append({
                    "user_id": user_id,
                    "display_name": info.get("display_name", user_id),
                    "connected_at": info.get("connected_at")
                })
        return users

    def add_history(self, sheet_key: str, entry: Dict):
        """添加修改历史记录"""
        if sheet_key not in self.edit_history:
            self.edit_history[sheet_key] = []

        # 添加时间戳
        entry["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 在列表头部插入新记录
        self.edit_history[sheet_key].insert(0, entry)

        # 限制历史记录条数
        if len(self.edit_history[sheet_key]) > self.max_history:
            self.edit_history[sheet_key] = self.edit_history[sheet_key][:self.max_history]

    def get_history(self, sheet_key: str, count: int = 20) -> List[Dict]:
        """获取修改历史记录"""
        if sheet_key not in self.edit_history:
            return []
        return self.edit_history[sheet_key][:count]

    async def send_personal(self, websocket: WebSocket, message: dict):
        """发送消息给单个用户"""
        try:
            await websocket.send_json(message)
        except Exception as e:
            print(f"发送消息失败: {e}")

    async def broadcast_to_sheet(self, sheet_key: str, message: dict, exclude: Optional[str] = None):
        """广播消息给某表格的所有用户"""
        if sheet_key not in self.active_connections:
            return

        disconnected = []
        for user_id, websocket in self.active_connections[sheet_key].items():
            if exclude and user_id == exclude:
                continue
            try:
                await websocket.send_json(message)
            except Exception as e:
                print(f"广播失败 {user_id}: {e}")
                disconnected.append(user_id)

        # 清理断开的连接
        for user_id in disconnected:
            self.disconnect(sheet_key, user_id)

    async def handle_cell_update(self, sheet_key: str, user_id: str, data: dict):
        """处理单元格更新"""
        row = data.get("row")
        col = data.get("col")
        value = data.get("value")
        style = data.get("style")

        display_name = self.user_info.get(user_id, {}).get("display_name", user_id)

        # 判断是内容修改还是格式修改
        action_desc = "修改单元格"
        value_display = str(value)[:50] if value else ""

        if style and isinstance(style, dict) and len(style) > 0:
            # 有样式变化，生成样式描述
            style_parts = []
            if style.get("bl"): style_parts.append("加粗")
            if style.get("it"): style_parts.append("斜体")
            if style.get("ul"): style_parts.append("下划线")
            if style.get("st"): style_parts.append("删除线")
            if style.get("bg"): style_parts.append("背景色")
            if style.get("cl"): style_parts.append("字体色")
            if style.get("bd"): style_parts.append("边框")
            if style.get("ht") or style.get("vt"): style_parts.append("对齐")
            if style.get("fs"): style_parts.append(f"字号{style['fs']}")

            if style_parts:
                action_desc = "修改格式(" + ",".join(style_parts) + ")"
                # 如果有值变化，同时标注
                if value is not None and str(value).strip():
                    action_desc = "修改内容和格式"

        # 添加到修改历史
        self.add_history(sheet_key, {
            "user": display_name,
            "action": action_desc,
            "cell": f"{chr(65 + col)}{row + 1}",
            "value": value_display
        })

        # 记录用户操作日志
        log_user_action(
            user_id=user_id,
            display_name=display_name,
            sheet_key=sheet_key,
            action_type="cell_update",
            details={"row": row, "col": col, "value": value, "style": style}
        )

        # 立即保存到Excel文件
        file_path = self.sheet_paths.get(sheet_key)
        if file_path:
            try:
                update_cell(file_path, row, col, value, style)
            except Exception as e:
                print(f"保存单元格失败: {e}")

        # 广播给其他用户（包含历史记录）
        await self.broadcast_to_sheet(sheet_key, {
            "type": "cell_update",
            "row": row,
            "col": col,
            "value": value,
            "style": style,
            "user_id": user_id,
            "display_name": display_name
        }, exclude=user_id)

        # 广播历史记录更新给所有用户
        await self.broadcast_to_sheet(sheet_key, {
            "type": "history_update",
            "history": self.get_history(sheet_key, 20)
        })

    async def handle_batch_update(self, sheet_key: str, user_id: str, updates: List[dict]):
        """处理批量单元格更新"""
        display_name = self.user_info.get(user_id, {}).get("display_name", user_id)

        # 记录用户操作日志
        log_user_action(
            user_id=user_id,
            display_name=display_name,
            sheet_key=sheet_key,
            action_type="batch_update",
            details={"updates": updates}
        )

        file_path = self.sheet_paths.get(sheet_key)
        if file_path:
            try:
                batch_update_cells(file_path, updates)
            except Exception as e:
                print(f"批量保存失败: {e}")

        # 广播给其他用户
        await self.broadcast_to_sheet(sheet_key, {
            "type": "batch_update",
            "updates": updates,
            "user_id": user_id,
            "display_name": display_name
        }, exclude=user_id)

    async def handle_cursor_move(self, sheet_key: str, user_id: str, data: dict):
        """处理光标移动（用于显示其他用户的选择区域）"""
        display_name = self.user_info.get(user_id, {}).get("display_name", user_id)
        await self.broadcast_to_sheet(sheet_key, {
            "type": "cursor_move",
            "row": data.get("row"),
            "col": data.get("col"),
            "user_id": user_id,
            "display_name": display_name
        }, exclude=user_id)

    async def handle_selection_change(self, sheet_key: str, user_id: str, data: dict):
        """处理选区变化"""
        display_name = self.user_info.get(user_id, {}).get("display_name", user_id)
        await self.broadcast_to_sheet(sheet_key, {
            "type": "selection_change",
            "selection": data.get("selection"),
            "user_id": user_id,
            "display_name": display_name
        }, exclude=user_id)

    async def handle_dimension_update(self, sheet_key: str, user_id: str, data: dict):
        """处理列宽行高更新"""
        col_widths = data.get("col_widths")
        row_heights = data.get("row_heights")

        display_name = self.user_info.get(user_id, {}).get("display_name", user_id)

        # 记录用户操作日志
        log_user_action(
            user_id=user_id,
            display_name=display_name,
            sheet_key=sheet_key,
            action_type="dimension_update",
            details={"col_widths": col_widths, "row_heights": row_heights}
        )

        # 保存到Excel文件
        file_path = self.sheet_paths.get(sheet_key)
        if file_path:
            try:
                batch_update_dimensions(file_path, col_widths, row_heights)
            except Exception as e:
                print(f"保存列宽行高失败: {e}")

        # 广播给其他用户
        await self.broadcast_to_sheet(sheet_key, {
            "type": "dimension_update",
            "col_widths": col_widths,
            "row_heights": row_heights,
            "user_id": user_id,
            "display_name": display_name
        }, exclude=user_id)

    async def process_message(self, sheet_key: str, user_id: str, message: str):
        """处理收到的WebSocket消息"""
        try:
            data = json.loads(message)
            msg_type = data.get("type")

            if msg_type == "cell_update":
                await self.handle_cell_update(sheet_key, user_id, data)
            elif msg_type == "batch_update":
                await self.handle_batch_update(sheet_key, user_id, data.get("updates", []))
            elif msg_type == "cursor_move":
                await self.handle_cursor_move(sheet_key, user_id, data)
            elif msg_type == "selection_change":
                await self.handle_selection_change(sheet_key, user_id, data)
            elif msg_type == "dimension_update":
                await self.handle_dimension_update(sheet_key, user_id, data)
            elif msg_type == "ping":
                # 心跳响应
                ws = self.active_connections.get(sheet_key, {}).get(user_id)
                if ws:
                    await self.send_personal(ws, {"type": "pong"})
            else:
                print(f"未知消息类型: {msg_type}")

        except json.JSONDecodeError as e:
            print(f"JSON解析失败: {e}")
        except Exception as e:
            print(f"消息处理失败: {e}")


# 全局连接管理器实例
manager = ConnectionManager()
