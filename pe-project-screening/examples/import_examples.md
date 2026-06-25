# 导入示例

```powershell
python scripts/init_db.py

python scripts/import_excel.py `
  "D:\wechatfile\xwechat_files\wxid_212yrh4z0oft22_bea2\msg\file\2026-06\初筛项目信息表20260621(1).xlsx"
```

导入模式固定为全量更新：

- 清空项目主表、初筛历史和旧导入异常；
- 保留操作日志；
- 从 Excel 重新建立项目及初始初筛记录；
- 同名项目逐行入库，不去重、不合并；
- 无法识别的字段保留原文并记录异常。
