from typing import List, Dict, Any, Optional
from datetime import datetime
from .connection import DatabaseConnection

class BaseModel:
    def __init__(self):
        self.db = DatabaseConnection()

class User(BaseModel):
    def create(self, username: str, email: str, password_hash: str) -> int:
        query = """
        INSERT INTO users (username, email, password_hash)
        VALUES (%s, %s, %s)
        """
        self.db.execute_update(query, (username, email, password_hash))
        user = self.get_by_username(username)
        user_id = user['id'] if user else None
        OperationLogger.log(user_id, 'create', 'users', user_id, f"username={username}")
        return user_id

    def get_by_id(self, user_id: int) -> Optional[Dict]:
        query = "SELECT * FROM users WHERE id = %s"
        result = self.db.execute_query(query, (user_id,))
        return result[0] if result else None

    def get_by_username(self, username: str) -> Optional[Dict]:
        query = "SELECT * FROM users WHERE username = %s"
        result = self.db.execute_query(query, (username,))
        return result[0] if result else None

class DocTask(BaseModel):
    def create(self, user_id: int, task_type: str, task_name: str, 
               input_path: str, output_path: str, template_id: Optional[int] = None) -> int:
        query = """
        INSERT INTO doc_tasks (user_id, task_type, task_name, input_path, output_path, template_id, start_time)
        VALUES (%s, %s, %s, %s, %s, %s, NOW())
        """
        self.db.execute_update(query, (user_id, task_type, task_name, input_path, output_path, template_id))
        # 获取新任务ID
        query_id = "SELECT id FROM doc_tasks WHERE user_id=%s AND task_name=%s ORDER BY created_at DESC LIMIT 1"
        result = self.db.execute_query(query_id, (user_id, task_name))
        task_id = result[0]['id'] if result else None
        OperationLogger.log(user_id, 'create', 'doc_tasks', task_id, f"user_id={user_id}, task_type={task_type}, task_name={task_name}")
        return task_id

    def update_status(self, task_id: int, status: str, message: Optional[str] = None, user_id: Optional[int] = None) -> bool:
        queries = [
            ("UPDATE doc_tasks SET status = %s, end_time = CASE WHEN %s IN ('success', 'failed', 'cancelled') THEN NOW() ELSE end_time END WHERE id = %s",
             (status, status, task_id)),
            ("INSERT INTO task_logs (task_id, log_type, message) VALUES (%s, 'info', %s)",
             (task_id, f"任务状态变更为 {status}" + (f": {message}" if message else "")))
        ]
        result = self.db.execute_transaction(queries)
        OperationLogger.log(user_id, 'update_status', 'doc_tasks', task_id, f"status={status}")
        return result

    def get_task_files(self, task_id: int) -> List[Dict]:
        query = "SELECT * FROM doc_files WHERE task_id = %s"
        result = self.db.execute_query(query, (task_id,))
        return result

class FormatTemplate(BaseModel):
    def create(self, user_id: int, name: str, config: Any, 
               description: Optional[str] = None, is_public: bool = False) -> int:
        query = """
        INSERT INTO format_templates (user_id, name, description, is_public, config)
        VALUES (%s, %s, %s, %s, %s)
        """
        self.db.execute_update(query, (user_id, name, description, is_public, config))
        # 获取新模板ID
        query_id = "SELECT id FROM format_templates WHERE user_id=%s AND name=%s ORDER BY created_at DESC LIMIT 1"
        result = self.db.execute_query(query_id, (user_id, name))
        template_id = result[0]['id'] if result else None
        OperationLogger.log(user_id, 'create', 'format_templates', template_id, f"name={name}")
        return template_id

    def get_by_id(self, template_id: int) -> Optional[Dict]:
        query = "SELECT * FROM format_templates WHERE id = %s"
        result = self.db.execute_query(query, (template_id,))
        return result[0] if result else None

    def get_user_templates(self, user_id: int) -> List[Dict]:
        query = """
        SELECT * FROM format_templates 
        WHERE user_id = %s OR is_public = TRUE
        """
        result = self.db.execute_query(query, (user_id,))
        return result

    def delete(self, template_id: int) -> bool:
        query = "DELETE FROM format_templates WHERE id = %s"
        result = self.db.execute_update(query, (template_id,))
        # 日志可选，需传user_id
        return result

class TaskLog(BaseModel):
    def add_log(self, task_id: int, log_type: str, message: str, user_id: Optional[int] = None) -> int:
        query = """
        INSERT INTO task_logs (task_id, log_type, message)
        VALUES (%s, %s, %s)
        """
        result = self.db.execute_update(query, (task_id, log_type, message))
        OperationLogger.log(user_id, 'add_log', 'task_logs', result, message)
        return result

    def get_task_logs(self, task_id: int) -> List[Dict]:
        query = "SELECT * FROM task_logs WHERE task_id = %s ORDER BY created_at DESC"
        result = self.db.execute_query(query, (task_id,))
        return result

class PerformanceLog(BaseModel):
    def add_log(self, task_id: int, operation: str, duration_ms: int, 
                memory_usage_mb: Optional[float] = None, user_id: Optional[int] = None) -> int:
        query = """
        INSERT INTO performance_logs (task_id, operation, start_time, end_time, duration_ms, memory_usage_mb)
        VALUES (%s, %s, NOW(), NOW(), %s, %s)
        """
        result = self.db.execute_update(query, (task_id, operation, duration_ms, memory_usage_mb))
        OperationLogger.log(user_id, 'add_log', 'performance_logs', result, f"operation={operation}, duration_ms={duration_ms}, memory_usage_mb={memory_usage_mb}")
        return result

    def get_task_performance(self, task_id: int) -> List[Dict]:
        query = "SELECT * FROM performance_logs WHERE task_id = %s ORDER BY created_at DESC"
        result = self.db.execute_query(query, (task_id,))
        return result

class OperationLogger:
    @staticmethod
    def log(user_id, action, table, record_id=None, detail=None):
        from database.connection import DatabaseConnection
        db = DatabaseConnection()
        query = """
            INSERT INTO system_logs (user_id, log_type, module, message, created_at)
            VALUES (%s, %s, %s, %s, %s)
        """
        msg = f"{action} {table} record_id={record_id or ''} {detail or ''}"
        db.execute_update(query, (user_id, 'info', table, msg, datetime.now())) 