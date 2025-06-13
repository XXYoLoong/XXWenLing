# 小小文灵（XXWenLing）

这是一个功能强大的文档处理工具集，包含文档合并和格式规范化处理功能。

## 功能特点

### 1. 文档合并功能
- 支持.doc和.docx格式的文档
- 图形用户界面，操作简单直观
- 自动在文档之间添加分页符
- 显示处理进度
- 保留原文档的段落和表格格式

### 2. 文档格式规范化处理功能
- 支持批量处理.docx文档
- 智能识别文档结构（标题、副标题、正文）
- 自定义格式模板
- 支持多种识别规则（位置、关键词、正则表达式、长度）
- 可配置的格式设置（字体、字号、对齐方式、缩进、行距等）
- 支持覆盖原文件或另存为新文件
- 详细的处理日志

## 安装要求

- Python 3.6或更高版本
- Microsoft Word（用于处理.doc文件）

## 安装步骤

1. 安装所需的Python包：
```bash
pip install -r requirements.txt
```

2. 运行程序：
```bash
python ui_main.py
```

## 使用说明

### 文档合并功能
1. 点击"选择文档目录"按钮选择包含要合并的文档的文件夹
2. 点击"合并文档"按钮开始合并过程
3. 在弹出的对话框中输入合并后的文件名
4. 等待处理完成，合并后的文档将保存在所选目录中

### 文档格式规范化处理功能
1. 点击"选择文档目录"按钮选择包含要处理的文档的文件夹
2. 选择或创建格式模板
3. 配置处理选项（是否覆盖原文件）
4. 点击"开始处理"按钮开始格式化过程
5. 等待处理完成，查看处理日志

## 注意事项

- 确保有足够的磁盘空间
- 处理大量文档时可能需要一些时间
- 建议在合并前备份原始文档
- 格式化处理前请确保文档未被其他程序占用
- 建议先使用小批量文档测试格式模板效果 

## 数据库系统介绍

本项目集成了完整的MySQL数据库系统，支持多用户、多任务、规则模板、日志等多维度数据管理。数据库设计与实现亮点如下：

### 1. 数据库功能
- 用户管理：支持用户注册、登录、配置管理。
- 任务管理：记录文档合并、格式化等批量处理任务及其状态。
- 文件管理：追踪每个任务下的所有文档文件。
- 模板管理：支持自定义格式模板的增删改查。
- 日志系统：详细记录每一次操作、任务进度、系统异常和性能数据。

### 2. 表结构及功能说明

| 表名                | 功能简介                                                                 |
|---------------------|--------------------------------------------------------------------------|
| users               | 用户信息表，存储所有注册用户的基本信息（用户名、邮箱、密码哈希等）         |
| user_settings       | 用户个性化配置表，记录每个用户的默认输出目录、默认模板、自动备份等设置      |
| doc_tasks           | 文档处理任务表，记录每一次文档合并/格式化等批量任务的详细信息和状态         |
| doc_files           | 文档文件表，记录每个任务下所有被处理的文档文件的明细（文件名、路径、状态等）|
| format_templates    | 格式模板表，存储用户自定义的文档格式化模板及其配置（支持公开/私有）         |
| system_logs         | 系统日志表，记录所有关键操作、异常、系统事件等，便于审计和问题追踪           |
| task_logs           | 任务日志表，记录每个任务的进度、警告、错误等详细日志                       |
| performance_logs    | 性能日志表，记录任务执行过程中的性能数据（耗时、内存等）                     |

- 这些表共同构成了完整的用户、任务、规则、日志、性能等多维度数据管理体系。
- 你可以通过SQL或代码随时查询、分析、追踪每一项数据和操作。

### 3. 日志机制
- 所有关键数据操作（如模板、任务、规则、日志等）均自动写入`system_logs`表。
- 日志内容包括：操作用户、操作类型（如create/update/delete）、目标表、记录ID、详细内容、时间戳等。
- 便于后续审计、问题追踪和数据恢复。

### 4. 代码集成方式
- 所有数据库操作均通过`database/models.py`中的模型类进行。
- 每个模型方法（如`create_user`、`create_task`、`update_status`、`create_template`等）在执行时自动调用`OperationLogger.log`写入操作日志。

#### 代码示例：
```python
from database.models import User, OperationLogger

# 创建用户并自动记录日志
user_id = User.create_user('testuser', 'test@example.com', 'hashed_pwd')
# 日志会自动写入system_logs表

# 更新任务状态
DocTask.update_status(task_id, 'success')
# 日志会自动写入system_logs表
```

### 5. 数据库初始化与自动同步
- 程序启动时自动调用`init_database()`，如未初始化则自动建库建表。
- 所有数据操作、规则录入、任务、日志等均实时同步到数据库。

### 6. 性能与安全
- 关键字段均有索引，支持高并发查询。
- 所有外键均有级联约束，保证数据一致性。
- 日志表支持追溯所有历史操作。

---
如需自定义数据库配置，请修改`config/database.py`。
如需扩展日志或表结构，请参考`database/models.py`和`database/init_db.py`。

## 数据库系统详细介绍

本项目数据库系统不仅支持常规的数据存储与管理，还具备完善的日志、事务、性能与问题追踪能力。以下为各核心环节的详细说明及操作记录查看方法：

### 1. 核心表创建
- 所有核心表（如`users`、`doc_tasks`、`format_templates`、`system_logs`等）在`database/init_db.py`中定义。
- 启动主程序或运行`python -m database.init_db`会自动建表。
- **查看表结构SQL**：
  ```sql
  SHOW CREATE TABLE users;
  SHOW CREATE TABLE doc_tasks;
  SHOW CREATE TABLE format_templates;
  SHOW CREATE TABLE system_logs;
  ```
- **查看表创建日志**：
  ```sql
  SELECT * FROM information_schema.tables WHERE table_schema = 'xxwenling';
  ```

### 2. 表操作（增删改查）
- 通过`database/models.py`的模型类进行数据操作。
- 例如：
  ```python
  from database.models import User, DocTask, FormatTemplate
  user_id = User.create_user('alice', 'alice@example.com', 'pwd')
  DocTask.create_task(user_id, 'merge', '合并任务', '/input', '/output', None)
  FormatTemplate.create_template(user_id, '标准模板', 'desc', True, '{}')
  ```
- **查看操作日志**：
  ```sql
  SELECT * FROM system_logs WHERE module IN ('users','doc_tasks','format_templates') ORDER BY created_at DESC;
  ```

### 3. 测试数据录入
- 可通过`examples/db_operations.py`批量插入测试数据。
- 运行：
  ```bash
  python -m examples.db_operations
  ```
- **查看录入结果**：
  ```sql
  SELECT * FROM users;
  SELECT * FROM doc_tasks;
  SELECT * FROM format_templates;
  SELECT * FROM system_logs WHERE log_type='info';
  ```

### 4. 核心存储过程
- 可在MySQL中自定义存储过程（如批量插入、批量日志写入等）。
- 示例：
  ```sql
  DELIMITER //
  CREATE PROCEDURE log_user_action(IN uid INT, IN msg TEXT)
  BEGIN
    INSERT INTO system_logs(user_id, log_type, module, message, created_at)
    VALUES(uid, 'info', 'users', msg, NOW());
  END //
  DELIMITER ;
  CALL log_user_action(1, '测试存储过程日志');
  ```
- **查看存储过程日志**：
  ```sql
  SELECT * FROM system_logs WHERE message LIKE '%存储过程日志%';
  ```

### 5. 事务测试案例
- 在`database/models.py`中有事务操作示例（如任务状态变更、批量插入等）。
- 你也可以在MySQL命令行测试：
  ```sql
  START TRANSACTION;
  UPDATE doc_tasks SET status='failed' WHERE id=1;
  ROLLBACK;
  SELECT status FROM doc_tasks WHERE id=1; -- 应为原状态
  ```
- **查看事务相关日志**：
  ```sql
  SELECT * FROM system_logs WHERE message LIKE '%update_status%';
  ```

### 6. 典型问题处理
- 例如外键约束失败、唯一索引冲突等。
- 代码中会捕获异常并写入`system_logs`。
- **查看异常日志**：
  ```sql
  SELECT * FROM system_logs WHERE log_type='error' OR message LIKE '%Exception%';
  ```

### 7. 性能优化对比
- 所有核心表关键字段均有索引。
- 可用如下SQL对比有无索引的查询性能：
  ```sql
  EXPLAIN SELECT * FROM doc_tasks WHERE user_id=1;
  -- 查看是否走索引
  SHOW INDEX FROM doc_tasks;
  ```
- **查看性能日志**：
  ```sql
  SELECT * FROM performance_logs ORDER BY created_at DESC;
  ```

### 8. 问题处理与多类型日志应用
- 日志类型包括：info、warning、error、debug、progress、performance等。
- 代码中所有操作均自动写入`system_logs`，并可根据类型灵活筛查。
- **查看不同类型日志**：
  ```sql
  SELECT * FROM system_logs WHERE log_type='info';
  SELECT * FROM system_logs WHERE log_type='error';
  SELECT * FROM system_logs WHERE log_type='debug';
  SELECT * FROM task_logs WHERE log_type='progress';
  SELECT * FROM performance_logs;
  ```

---
如需进一步扩展数据库功能、日志类型或性能分析，请参考`database/models.py`和`database/init_db.py`，或联系开发者获取更多示例。

## 数据库写入测试流程

本节提供一套完整的数据库写入验证流程，帮助你确认各功能是否已成功写入数据库。

### 1. 测试模板写入
- **操作前查询：**
  ```sql
  SELECT * FROM format_templates;
  ```
- **程序操作：**
  - 启动程序，注册/登录。
  - 进入"文档格式化"界面，点击"新建模板"，填写信息并保存。
- **操作后查询：**
  ```sql
  SELECT * FROM format_templates ORDER BY created_at DESC;
  SELECT * FROM system_logs WHERE module='format_templates' ORDER BY created_at DESC;
  ```
- **预期结果：**
  - 新增模板数据出现在`format_templates`表。
  - 操作日志写入`system_logs`。

### 2. 测试合并任务写入
- **操作前查询：**
  ```sql
  SELECT * FROM doc_tasks WHERE task_type='merge' ORDER BY created_at DESC;
  SELECT * FROM task_logs ORDER BY created_at DESC;
  SELECT * FROM performance_logs ORDER BY created_at DESC;
  ```
- **程序操作：**
  - 在主界面进入"文档合并"，选择目录、输入文件名，点击"开始合成"。
- **操作后查询：**
  ```sql
  SELECT * FROM doc_tasks WHERE task_type='merge' ORDER BY created_at DESC;
  SELECT * FROM task_logs WHERE log_type IN ('progress','info','error') ORDER BY created_at DESC;
  SELECT * FROM performance_logs WHERE operation='merge' ORDER BY created_at DESC;
  ```
- **预期结果：**
  - 新增合并任务、进度日志、性能日志。

### 3. 测试格式化任务写入
- **操作前查询：**
  ```sql
  SELECT * FROM doc_tasks WHERE task_type='format' ORDER BY created_at DESC;
  SELECT * FROM task_logs ORDER BY created_at DESC;
  SELECT * FROM performance_logs ORDER BY created_at DESC;
  ```
- **程序操作：**
如需进一步扩展数据库功能、日志类型或性能分析，请参考`database/models.py`和`database/init_db.py`，或联系开发者获取更多示例。 