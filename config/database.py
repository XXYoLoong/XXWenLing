from typing import Dict
import os

# 数据库配置
DB_CONFIG: Dict = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': int(os.getenv('DB_PORT', 3306)),
    'user': os.getenv('DB_USER', 'root'),
    'password': os.getenv('DB_PASSWORD', '5211005jc'),
    'database': os.getenv('DB_NAME', 'xxwenling'),
    'charset': 'utf8mb4'
}

# 数据库连接池配置
POOL_CONFIG: Dict = {
    'pool_name': 'xxwenling_pool',
    'pool_size': 5,
    'pool_reset_session': True
} 