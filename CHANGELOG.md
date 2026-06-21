# 更新日志

## Unreleased

### Added

- 新增 `模板类型` 字段，允许操作人员在飞书或 Excel 的本次记录中手动选择模板。
- 新增 MIT License。
- 新增 `tests/test_generate_docs.py`，使用 Python 标准库 `unittest` 验证核心规则。
- 演示 Excel 和演示飞书 JSON 增加模板类型示例。

### Changed

- 模板选择优先级调整为：`本次记录的模板类型 > 知识库里的部门`。
- 生成指纹纳入 `模板类型`，模板选择变化后会触发重新生成。
- README 补充公开发布说明、本地检查命令、测试命令和许可证说明。
- Word 模板元数据统一清理为公开项目贡献者信息，避免暴露本机编辑用户名。

### Privacy

- 公开仓库只应提交演示库 `knowledge_base.example.db`，不要提交真实工作库 `knowledge_base.db`。
- `.env`、`generated.json`、生产 Excel、输出函件、飞书二维码和内部模板应只保存在本地。
- 当前演示数据使用 `张三/李四/王五` 与 `E00000001/E00000002` 这类占位值。
