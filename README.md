# 采购自动化工作台（装潢 / 空调 / IC卡）

面向采购 Excel 的桌面自动化：生成透视表、按 PO 查询和下载附件、分阶段回填，并匹配 PM Tracking 数据。

## 能做什么

- 从采购记录生成装潢、空调或 IC卡透视表，并补充业务字段。
- 在 Edge 中按 PO 查询、下载附件，自动按是否存在真正的 FLD 附件分类。
- 将“网页查询下载”和“本地解析回填”分开执行；查询中断后可继续回填。
- 对缺失的 PO 仅补查缺失项，不必重跑全部 PO。
- 匹配 NI PM Tracking，回填 Price、PlanCost、总 Saving 和订单差异。
- 空调和 IC卡额外支持老价格、新价格、Saving、CHECK 的批量计算；FLD 项目按 PO 分组合并显示。

## 开始使用

1. 安装 Python 3.8+，在项目目录运行：

   ```powershell
   pip install -r requirements.txt
   playwright install msedge
   python 装潢透视表工具.py
   ```

2. 将业务 Excel 文件放在项目目录（这些文件不会提交到 Git）。首次网页查询时，在打开的 Edge 中完成公司 SSO 登录。
3. 按工作台中的步骤执行。通常顺序为：生成透视表 → 添加扩展列 → 查询下载 → 回填 → PM 匹配；空调或 IC卡最后再执行价格填充。

详细输入文件、各按钮职责和维护方式见 [项目文档.md](项目文档.md)。

## 验证

```powershell
py -3.8 -m unittest discover -s tests -v
py -3.8 -m py_compile 装潢透视表工具.py web_query.py
```

## 数据与隐私

仓库只保存代码、测试、依赖和通用说明。登录状态、浏览器配置、采购 Excel、附件、运行日志与诊断报告均被 `.gitignore` 排除；请勿提交真实凭据或业务数据。
