import pymysql
import traceback

class MySQLService:
    """
    MySQL数据库服务类，提供数据库连接、查询和更新等功能
    """
    def __init__(self):
        """初始化MySQL服务类"""
        self.connection = None

    def connect(self, host='localhost', user='root', password='', database='', port=3306):
        """
        连接到MySQL数据库
        
        参数:
            host (str): 数据库主机地址，默认为localhost
            user (str): 数据库用户名，默认为root
            password (str): 数据库密码，默认为空
            database (str): 数据库名称，默认为空
            port (int): 数据库端口，默认为3306
        
        返回:
            bool: 连接是否成功
        """
        try:
            # 创建数据库连接
            self.connection = pymysql.connect(
                host=host,
                user=user,
                password=password,
                database=database,
                port=port,
                cursorclass=pymysql.cursors.DictCursor  # 使用字典游标，结果以字典形式返回
            )
            print(f"成功连接到数据库: {host}:{port}/{database}")
            return True
        except Exception as e:
            print(f"数据库连接失败: {e}")
            traceback.print_exc()
            return False

    def disconnect(self):
        """
        关闭数据库连接
        
        返回:
            bool: 断开连接是否成功
        """
        try:
            if self.connection and not self.connection._closed:
                self.connection.close()
                print("数据库连接已关闭")
            return True
        except Exception as e:
            print(f"关闭数据库连接失败: {e}")
            traceback.print_exc()
            return False

    def execute_query(self, sql, params=None):
        """
        执行SQL查询语句
        
        参数:
            sql (str): SQL查询语句
            params (tuple, list or dict): SQL参数，用于防止SQL注入
        
        返回:
            list: 查询结果列表，每个元素为字典
        """
        try:
            if not self.connection or self.connection._closed:
                raise Exception("数据库未连接")
            
            with self.connection.cursor() as cursor:
                cursor.execute(sql, params)
                result = cursor.fetchall()
                return result
        except Exception as e:
            print(f"SQL查询执行失败: {e}")
            traceback.print_exc()
            return []

    def execute_update(self, sql, params=None):
        """
        执行SQL更新语句（插入、更新、删除）
        
        参数:
            sql (str): SQL更新语句
            params (tuple, list or dict): SQL参数，用于防止SQL注入
        
        返回:
            int: 受影响的行数，执行失败返回-1
        """
        try:
            if not self.connection or self.connection._closed:
                raise Exception("数据库未连接")
            
            with self.connection.cursor() as cursor:
                affected_rows = cursor.execute(sql, params)
                self.connection.commit()  # 提交事务
                return affected_rows
        except Exception as e:
            print(f"SQL更新执行失败: {e}")
            traceback.print_exc()
            if self.connection and not self.connection._closed:
                self.connection.rollback()  # 回滚事务
            return -1

    def execute_many(self, sql, params_list):
        """
        批量执行SQL语句
        
        参数:
            sql (str): SQL语句
            params_list (list): 参数列表，每个元素为tuple、list或dict
        
        返回:
            int: 受影响的行数，执行失败返回-1
        """
        try:
            if not self.connection or self.connection._closed:
                raise Exception("数据库未连接")
            
            with self.connection.cursor() as cursor:
                affected_rows = cursor.executemany(sql, params_list)
                self.connection.commit()  # 提交事务
                return affected_rows
        except Exception as e:
            print(f"批量SQL执行失败: {e}")
            traceback.print_exc()
            if self.connection and not self.connection._closed:
                self.connection.rollback()  # 回滚事务
            return -1

    def create_table(self, table_name, fields_info, if_not_exists=True, engine='InnoDB', charset='utf8mb4'):
        """
        创建MySQL数据表
        
        参数:
            table_name (str): 表名
            fields_info (list): 字段信息列表，每个元素为包含'name'(字段名), 'type'(字段类型), 'comment'(注释)的字典
                               可选字段: 'primary_key'(是否主键), 'auto_increment'(是否自增), 'not_null'(是否非空)
            if_not_exists (bool): 是否添加IF NOT EXISTS子句，默认为True
            engine (str): 存储引擎，默认为InnoDB
            charset (str): 字符集，默认为utf8mb4
        
        返回:
            bool: 创建表是否成功
        """
        try:
            if not self.connection or self.connection._closed:
                raise Exception("数据库未连接")
            
            # 构建CREATE TABLE语句
            create_table_sql = f"CREATE TABLE {'IF NOT EXISTS ' if if_not_exists else ''}`{table_name}` (\n"
            
            # 构建字段定义部分
            fields_definitions = []
            primary_keys = []
            
            for field_info in fields_info:
                field_name = field_info['name']
                field_type = field_info['type']
                field_comment = field_info.get('comment', '')
                
                field_def = f"  `{field_name}` {field_type}"
                
                # 添加NOT NULL约束
                if field_info.get('not_null', False):
                    field_def += " NOT NULL"
                
                # 添加AUTO_INCREMENT属性
                if field_info.get('auto_increment', False):
                    field_def += " AUTO_INCREMENT"
                
                # 添加COMMENT
                if field_comment:
                    field_def += f" COMMENT '{field_comment}'"
                
                fields_definitions.append(field_def)
                
                # 收集主键字段
                if field_info.get('primary_key', False):
                    primary_keys.append(field_name)
            
            # 添加主键约束
            if primary_keys:
                primary_key_str = f"  PRIMARY KEY ({', '.join([f'`{pk}`' for pk in primary_keys])})"
                fields_definitions.append(primary_key_str)
            
            # 组合完整的CREATE TABLE语句
            create_table_sql += ',\n'.join(fields_definitions)
            create_table_sql += f"\n) ENGINE={engine} DEFAULT CHARSET={charset};"
            
            # 执行创建表语句
            with self.connection.cursor() as cursor:
                cursor.execute(create_table_sql)
                self.connection.commit()
                print(f"数据表 `{table_name}` 创建成功")
                return True
        except Exception as e:
            print(f"创建数据表 `{table_name}` 失败: {e}")
            traceback.print_exc()
            if self.connection and not self.connection._closed:
                self.connection.rollback()
            return False

    @staticmethod
    def create_example_table():
        """
        创建示例数据表的静态方法，展示如何使用create_table方法
        
        返回:
            bool: 创建表是否成功
        """
        # 创建MySQLService实例
        db_service = MySQLService()
        
        # 连接到数据库（实际使用时请修改为您的数据库信息）
        if not db_service.connect(
            host='localhost',
            user='root',
            password='your_password',  # 替换为您的数据库密码
            database='excel_import_db',  # 替换为您的数据库名称
            port=3306
        ):
            return False
        
        # 定义用户表的字段信息
        user_fields = [
            {
                'name': 'id',
                'type': 'INT',
                'comment': '用户ID，主键',
                'primary_key': True,
                'auto_increment': True,
                'not_null': True
            },
            {
                'name': 'username',
                'type': 'VARCHAR(50)',
                'comment': '用户名',
                'not_null': True
            },
            {
                'name': 'email',
                'type': 'VARCHAR(100)',
                'comment': '用户邮箱',
                'not_null': True
            },
            {
                'name': 'age',
                'type': 'INT',
                'comment': '用户年龄',
                'not_null': False
            },
            {
                'name': 'created_at',
                'type': 'DATETIME',
                'comment': '创建时间',
                'not_null': True,
                'default': 'CURRENT_TIMESTAMP'
            }
        ]
        
        # 创建用户表
        result = db_service.create_table('users', user_fields)
        
        # 断开数据库连接
        db_service.disconnect()
        
        return result

# 创建一个全局的MySQL服务实例，方便其他模块直接导入使用
mysql_service = MySQLService()