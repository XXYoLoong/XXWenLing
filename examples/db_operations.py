from database.models import User, DocTask, FormatTemplate, TaskLog, PerformanceLog
import json
import time

def example_operations():
    # 创建用户
    user = User()
    user_id = user.create(
        username="test_user",
        email="test@example.com",
        password_hash="hashed_password"
    )
    print(f"创建用户成功，ID: {user_id}")

    # 创建模板
    template = FormatTemplate()
    template_config = {
        "title": {
            "font": "宋体",
            "size": 16,
            "alignment": "center"
        },
        "body": {
            "font": "仿宋",
            "size": 12,
            "alignment": "left",
            "indent": 2
        }
    }
    template_id = template.create(
        user_id=user_id,
        name="标准公文模板",
        config=json.dumps(template_config),
        description="适用于标准公文格式",
        is_public=True
    )
    print(f"创建模板成功，ID: {template_id}")

    # 创建任务
    task = DocTask()
    task_id = task.create(
        user_id=user_id,
        task_type="format",
        task_name="测试格式化任务",
        input_path="/input/docs",
        output_path="/output/docs",
        template_id=template_id
    )
    print(f"创建任务成功，ID: {task_id}")

    # 添加任务日志
    task_log = TaskLog()
    task_log.add_log(
        task_id=task_id,
        log_type="info",
        message="任务开始处理"
    )

    # 模拟任务处理
    time.sleep(1)  # 模拟处理时间

    # 记录性能日志
    perf_log = PerformanceLog()
    perf_log.add_log(
        task_id=task_id,
        operation="format_document",
        duration_ms=1000,
        memory_usage_mb=50.5
    )

    # 更新任务状态
    task.update_status(
        task_id=task_id,
        status="success",
        message="处理完成"
    )

    # 查询任务信息
    task_files = task.get_task_files(task_id)
    print(f"任务文件数: {len(task_files)}")

    # 查询用户模板
    user_templates = template.get_user_templates(user_id)
    print(f"用户模板数: {len(user_templates)}")

if __name__ == "__main__":
    example_operations() 