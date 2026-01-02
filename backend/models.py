"""数据模型定义"""
from pydantic import BaseModel
from typing import Optional, Any, List
from datetime import datetime


class SheetKey(BaseModel):
    """表格密钥模型"""
    key: str
    name: str
    file_path: str
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None


class SheetKeyCreate(BaseModel):
    """创建密钥请求"""
    name: str
    key: Optional[str] = None  # 可选，不提供则自动生成


class AuthRequest(BaseModel):
    """认证请求"""
    key: str


class AuthResponse(BaseModel):
    """认证响应"""
    success: bool
    message: str
    sheet_name: Optional[str] = None
    token: Optional[str] = None


class CellUpdate(BaseModel):
    """单元格更新消息"""
    type: str = "cell_update"
    sheet_id: str
    row: int
    col: int
    value: Any
    style: Optional[dict] = None
    user_id: str


class UserInfo(BaseModel):
    """用户信息"""
    user_id: str
    ip_address: str
    mac_address: str


class WSMessage(BaseModel):
    """WebSocket消息基类"""
    type: str
    data: Any


class OnlineUser(BaseModel):
    """在线用户"""
    user_id: str
    ip_address: str
    mac_address: str
    connected_at: Optional[datetime] = None


class SheetData(BaseModel):
    """表格数据"""
    id: str
    name: str
    data: List[List[Any]]  # 二维数组表示单元格数据
    styles: Optional[dict] = None  # 样式信息
    merges: Optional[List[dict]] = None  # 合并单元格信息
    col_widths: Optional[List[int]] = None  # 列宽
    row_heights: Optional[List[int]] = None  # 行高
