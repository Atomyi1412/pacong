import inspect
import baseopensdk.api.base.v1 as base_v1
import pkgutil

from baseopensdk.api.base.v1.model.app_table_field import AppTableField
from baseopensdk.api.base.v1.model.create_app_table_field_request import CreateAppTableFieldRequest

print("AppTableField.builder() methods:")
for name, member in inspect.getmembers(AppTableField.builder()):
    if not name.startswith("__"):
        print(name)

print("\nCreateAppTableFieldRequest.builder() methods:")
for name, member in inspect.getmembers(CreateAppTableFieldRequest.builder()):
    if not name.startswith("__"):
        print(name)
        
from baseopensdk.api.base.v1.model.list_app_table_field_request import ListAppTableFieldRequest
print("\nListAppTableFieldRequest.builder() methods:")
for name, member in inspect.getmembers(ListAppTableFieldRequest.builder()):
    if not name.startswith("__"):
        print(name)
