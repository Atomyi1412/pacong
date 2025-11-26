import os

try:
    from baseopensdk.api.base.v1.model.app_table_field import AppTableField
    from baseopensdk.api.base.v1.model.create_app_table_field_request import CreateAppTableFieldRequest
    from baseopensdk.api.base.v1.model.list_app_table_field_request import ListAppTableFieldRequest
    SDK_AVAILABLE = True
except ImportError:
    SDK_AVAILABLE = False

def ensure_fields_exist(client, app_token, table_id, sample_data):
    """
    检查多维表中是否存在 sample_data 中的字段，如果不存在则自动创建。
    返回现有字段的名称和类型字典 {name: type}。
    """
    if not SDK_AVAILABLE:
        print("BaseOpenSDK 未安装，无法执行字段检查。")
        return {}

    try:
        # 1. 获取现有字段列表
        list_req = ListAppTableFieldRequest.builder() \
            .app_token(app_token) \
            .table_id(table_id) \
            .build()
        
        resp = client.base.v1.app_table_field.list(list_req)
        existing_fields = {item.field_name: item.type for item in resp.data.items}
        
        # 2. 遍历样本数据，检查字段是否存在
        for field_name, field_value in sample_data.items():
            if field_name not in existing_fields:
                print(f"检测到字段 '{field_name}' 缺失，正在尝试自动创建...")
                
                # 根据值推断字段类型
                # 默认文本类型 (1)
                field_type = 1 
                
                if isinstance(field_value, int) or isinstance(field_value, float):
                    # 如果是时间戳（通常很大），或者明确是日期
                    # 这里简单判断，如果字段名包含"时间"，则设为日期 (5)
                    if "时间" in field_name:
                        field_type = 5 # 日期
                    else:
                        field_type = 2 # 数字
                elif isinstance(field_value, dict) and "link" in field_value:
                    field_type = 15 # 超链接
                
                # 创建字段请求
                create_req = CreateAppTableFieldRequest.builder() \
                    .app_token(app_token) \
                    .table_id(table_id) \
                    .request_body(
                        AppTableField.builder()
                        .field_name(field_name)
                        .type(field_type)
                        .build()
                    ) \
                    .build()
                
                try:
                    client.base.v1.app_table_field.create(create_req)
                    print(f"成功创建字段: {field_name} (类型: {field_type})")
                    existing_fields[field_name] = field_type # 更新本地缓存
                except Exception as e:
                    print(f"创建字段 '{field_name}' 失败: {e}")
        
        return existing_fields
                    
    except Exception as e:
        print(f"检查或创建字段时发生错误: {e}")
        return {}
