import mysql.connector
from mysql.connector import pooling
from typing import Optional
from config.database import DB_CONFIG, POOL_CONFIG

class DatabaseConnection:
    _instance = None
    _pool = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(DatabaseConnection, cls).__new__(cls)
        return cls._instance

    def __init__(self):
        if self._pool is None:
            try:
                self._pool = mysql.connector.pooling.MySQLConnectionPool(
                    **DB_CONFIG,
                    **POOL_CONFIG
                )
            except Exception as e:
                raise Exception(f"数据库连接池初始化失败: {str(e)}")

    def get_connection(self):
        """获取数据库连接"""
        try:
            return self._pool.get_connection()
        except Exception as e:
            raise Exception(f"获取数据库连接失败: {str(e)}")

    def execute_query(self, query: str, params: tuple = None) -> list:
        """执行查询操作"""
        conn = None
        cursor = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query, params or ())
            return cursor.fetchall()
        except Exception as e:
            raise Exception(f"查询执行失败: {str(e)}")
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def execute_update(self, query: str, params: tuple = None) -> int:
        """执行更新操作"""
        conn = None
        cursor = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(query, params or ())
            conn.commit()
            return cursor.rowcount
        except Exception as e:
            if conn:
                conn.rollback()
            raise Exception(f"更新执行失败: {str(e)}")
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def execute_transaction(self, queries: list) -> bool:
        """执行事务"""
        conn = None
        cursor = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            
            for query, params in queries:
                cursor.execute(query, params or ())
            
            conn.commit()
            return True
        except Exception as e:
            if conn:
                conn.rollback()
            raise Exception(f"事务执行失败: {str(e)}")
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close() 