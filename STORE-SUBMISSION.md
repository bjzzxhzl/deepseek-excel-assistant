# Microsoft Marketplace 提交清单

## 先解决正式托管

Office Web Add-in 的页面、脚本、样式和图标必须由可公开访问的 HTTPS 服务长期托管。当前清单依赖 GitHub Pages；GitHub 不可访问时，功能区图标可能只剩缓存，任务窗格无法可靠打开。

上线前建议迁移到自有域名的高可用 HTTPS 托管（例如 Azure Static Web Apps、Azure Storage Static Website + CDN，或符合目标地区网络要求的等价服务），然后运行：

```powershell
node deploy.mjs https://你的正式域名/路径
```

不要把 API Key、测试账号密码或其他密钥放进静态文件。当前 BYOK 密钥仍应只保存在 Office WebView 的本地存储中。

## Partner Center 流程

1. 使用公司/组织工作账号创建或加入 Partner Center，并注册 Microsoft Marketplace 发布者；个人 Microsoft 账号不能代替公司工作账号。
2. 完成公司法定名称、地址、主要联系人和发布者验证，并接受 Microsoft Publisher Agreement。
3. 在 Partner Center 的 Microsoft Marketplace / Microsoft 365 与 Copilot 区域创建 Office Add-in 产品。
4. 上传通过验证的 `manifest.xml`，填写市场、语言、类别、说明、关键词、支持方式、隐私政策、使用条款和测试说明。
5. 上传商店图标和无个人信息的产品截图；商店图标必须与清单 `IconUrl` 对应的图标一致。
6. 提供审核人员可以实际完成核心流程的测试方法。若功能需要第三方 API Key，应提供受限额度的审核测试 Key 或清楚的无 Key 首次体验。
7. 运行 Office 清单验证与跨平台测试，完成 Partner Center 的 Review and publish，然后进入自动验证、人工认证、预览确认和发布阶段。

## 当前项目提交前的阻断项

- 将 GitHub Pages 换成面向目标用户稳定可访问的正式 HTTPS 域名。
- 建立独立、公开可访问的产品主页、支持页面、隐私政策和使用条款；当前 `SupportUrl` 只是加载项页面，不能作为合格支持站点。
- 产品已采用独立品牌名 **ExcelAI**；仍需确认第三方模型名称与相关素材的展示方式不会暗示官方背书，并在说明中准确描述兼容关系。
- 在 Excel Windows、Mac、Web 及清单声明支持的浏览器中验证全部核心流程。无法跨平台工作的功能必须调整清单要求或在市场说明中明确限制。
- 明确披露工作簿数据、对话、附件和 API Key 的处理方式；确认所有第三方模型供应商、跨境传输和日志策略与隐私政策一致。
- 准备审核测试账号/Key、操作步骤、预期结果和故障排查联系方式。

## 官方资料

- 发布 Office Add-in：https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource
- Partner Center 发布流程：https://learn.microsoft.com/en-us/partner-center/marketplace-offers/submit-to-appsource-via-partner-center
- 创建 Marketplace 发布者账号：https://learn.microsoft.com/en-us/partner-center/account-settings/create-account
- Marketplace 认证政策：https://learn.microsoft.com/en-us/legal/marketplace/certification-policies
- 商店列表与图标要求：https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings
- 集中部署自定义加载项：https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins?view=o365-worldwide
