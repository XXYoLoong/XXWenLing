import mysql.connector
from config.database import DB_CONFIG
import os

def create_index_if_not_exists(cursor, index_name, table_name, column_name):
    cursor.execute(f"SHOW INDEX FROM {table_name} WHERE Key_name = '{index_name}'")
    result = cursor.fetchone()
    if not result:
        cursor.execute(f"CREATE INDEX {index_name} ON {table_name}({column_name})")

def init_database():
    """初始化数据库"""
    # 创建数据库连接（不指定数据库名）
    conn = mysql.connector.connect(
        host=DB_CONFIG['host'],
        port=DB_CONFIG['port'],
        user=DB_CONFIG['user'],
        password=DB_CONFIG['password']
    )
    cursor = conn.cursor()

    try:
        # 创建数据库
        cursor.execute(f"CREATE DATABASE IF NOT EXISTS {DB_CONFIG['database']} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
        cursor.execute(f"USE {DB_CONFIG['database']}")

        # 创建用户表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INT AUTO_INCREMENT PRIMARY KEY COMMENT '用户ID',
            username VARCHAR(50) NOT NULL UNIQUE COMMENT '用户名',
            email VARCHAR(100) COMMENT '邮箱',
            password_hash VARCHAR(128) COMMENT '密码哈希',
            last_login DATETIME COMMENT '最后登录时间',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间'
        ) COMMENT '用户信息表'
        """)

        # 创建用户配置表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS user_settings (
            id INT AUTO_INCREMENT PRIMARY KEY COMMENT '配置ID',
            user_id INT NOT NULL COMMENT '用户ID',
            default_output_dir VARCHAR(255) COMMENT '默认输出目录',
            default_template_id INT COMMENT '默认模板ID',
            auto_backup BOOLEAN DEFAULT TRUE COMMENT '是否自动备份',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        ) COMMENT '用户配置表'
        """)

        # 创建文档任务表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS doc_tasks (
            id INT AUTO_INCREMENT PRIMARY KEY COMMENT '任务ID',
            user_id INT NOT NULL COMMENT '用户ID',
            task_type ENUM('merge', 'format') NOT NULL COMMENT '任务类型：合并/格式化',
            task_name VARCHAR(100) COMMENT '任务名称',
            status ENUM('pending', 'running', 'success', 'failed', 'cancelled') DEFAULT 'pending' COMMENT '任务状态',
            input_path VARCHAR(255) COMMENT '输入路径',
            output_path VARCHAR(255) COMMENT '输出路径',
            template_id INT COMMENT '使用的模板ID',
            total_files INT DEFAULT 0 COMMENT '总文件数',
            processed_files INT DEFAULT 0 COMMENT '已处理文件数',
            error_count INT DEFAULT 0 COMMENT '错误数',
            start_time DATETIME COMMENT '开始时间',
            end_time DATETIME COMMENT '结束时间',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        ) COMMENT '文档处理任务表'
        """)

        # 创建文档文件表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS doc_files (
            id INT AUTO_INCREMENT PRIMARY KEY COMMENT '文件ID',
            task_id INT NOT NULL COMMENT '所属任务ID',
            file_name VARCHAR(255) NOT NULL COMMENT '文件名',
            file_path VARCHAR(255) NOT NULL COMMENT '文件路径',
            file_size BIGINT COMMENT '文件大小(字节)',
            file_type ENUM('doc', 'docx') COMMENT '文件类型',
            status ENUM('pending', 'processing', 'success', 'failed') DEFAULT 'pending' COMMENT '处理状态',
            error_message TEXT COMMENT '错误信息',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (task_id) REFERENCES doc_tasks(id) ON DELETE CASCADE
        ) COMMENT '文档文件表'
        """)

        # 创建格式模板表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS format_templates (
            id INT AUTO_INCREMENT PRIMARY KEY COMMENT '模板ID',
            user_id INT NOT NULL COMMENT '创建用户ID',
            name VARCHAR(100) NOT NULL COMMENT '模板名称',
            description TEXT COMMENT '模板描述',
            is_public BOOLEAN DEFAULT FALSE COMMENT '是否公开',
            config JSON COMMENT '模板配置(JSON格式)',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        ) COMMENT '格式模板表'
        """)

        # 创建系统日志表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS system_logs (
            id INT AUTO_INCREMENT PRIMARY KEY COMMENT '日志ID',
            user_id INT COMMENT '用户ID',
            log_type ENUM('info', 'warning', 'error', 'debug') NOT NULL COMMENT '日志类型',
            module VARCHAR(50) COMMENT '模块名称',
            message TEXT NOT NULL COMMENT '日志消息',
            stack_trace TEXT COMMENT '堆栈信息',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
        ) COMMENT '系统日志表'
        """)

        # 创建任务日志表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS task_logs (
            id INT AUTO_INCREMENT PRIMARY KEY COMMENT '日志ID',
            task_id INT NOT NULL COMMENT '任务ID',
            log_type ENUM('info', 'warning', 'error', 'progress') NOT NULL COMMENT '日志类型',
            message TEXT NOT NULL COMMENT '日志消息',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (task_id) REFERENCES doc_tasks(id) ON DELETE CASCADE
        ) COMMENT '任务日志表'
        """)

        # 创建性能日志表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS performance_logs (
            id INT AUTO_INCREMENT PRIMARY KEY COMMENT '日志ID',
            task_id INT NOT NULL COMMENT '任务ID',
            operation VARCHAR(50) NOT NULL COMMENT '操作名称',
            start_time DATETIME NOT NULL COMMENT '开始时间',
            end_time DATETIME NOT NULL COMMENT '结束时间',
            duration_ms INT NOT NULL COMMENT '耗时(毫秒)',
            memory_usage_mb FLOAT COMMENT '内存使用(MB)',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (task_id) REFERENCES doc_tasks(id) ON DELETE CASCADE
        ) COMMENT '性能日志表'
        """)

        # 创建索引（用函数判断是否已存在）
        create_index_if_not_exists(cursor, 'idx_users_username', 'users', 'username')
        create_index_if_not_exists(cursor, 'idx_users_email', 'users', 'email')
        create_index_if_not_exists(cursor, 'idx_doc_tasks_user_id', 'doc_tasks', 'user_id')
        create_index_if_not_exists(cursor, 'idx_doc_tasks_status', 'doc_tasks', 'status')
        create_index_if_not_exists(cursor, 'idx_doc_tasks_type', 'doc_tasks', 'task_type')
        create_index_if_not_exists(cursor, 'idx_doc_tasks_created', 'doc_tasks', 'created_at')
        create_index_if_not_exists(cursor, 'idx_doc_files_task_id', 'doc_files', 'task_id')
        create_index_if_not_exists(cursor, 'idx_doc_files_status', 'doc_files', 'status')
        create_index_if_not_exists(cursor, 'idx_format_templates_user_id', 'format_templates', 'user_id')
        create_index_if_not_exists(cursor, 'idx_format_templates_public', 'format_templates', 'is_public')
        create_index_if_not_exists(cursor, 'idx_system_logs_type', 'system_logs', 'log_type')
        create_index_if_not_exists(cursor, 'idx_system_logs_created', 'system_logs', 'created_at')
        create_index_if_not_exists(cursor, 'idx_task_logs_task_id', 'task_logs', 'task_id')
        create_index_if_not_exists(cursor, 'idx_performance_logs_task_id', 'performance_logs', 'task_id')

        conn.commit()
        print("数据库初始化成功！")

    except Exception as e:
        conn.rollback()
        print(f"数据库初始化失败: {str(e)}")
        raise

    finally:
        cursor.close()
        conn.close()

if __name__ == "__main__":
    init_database() 