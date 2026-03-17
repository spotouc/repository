# Excel AI 数据清洗插件

在 Excel 中调用 SiliconFlow API 进行智能数据清洗的 Office Add-in。

## 部署到 GitHub Pages（推荐）

### 一键部署

[![Deploy to GitHub Pages](https://img.shields.io/badge/Deploy-GitHub%20Pages-blue)](https://pages.github.com/)

### 详细步骤

#### 1. 在 GitHub 创建仓库

```bash
# 在 GitHub 网站创建新仓库，例如：excel-ai-cleaner
# 不要勾选 README、.gitignore 或 license
```

#### 2. 初始化本地 Git 仓库并推送

```bash
cd /Users/ztzcl/cursor/ai-zhushou

# 初始化 Git
git init

# 添加所有文件
git add .

# 提交
git commit -m "Initial commit: Excel AI data cleaning plugin"

# 添加远程仓库（替换为你的用户名和仓库名）
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git

# 推送到 GitHub
git branch -M main
git push -u origin main
```

#### 3. 修改 manifest.xml

打开 `manifest.xml`，将所有 `YOUR_USERNAME` 和 `YOUR_REPO` 替换为你的实际用户名和仓库名：

- `YOUR_USERNAME` → 你的 GitHub 用户名
- `YOUR_REPO` → 你的仓库名（例如：excel-ai-cleaner）

例如，如果你的用户名是 `zhangsan`，仓库是 `excel-ai-cleaner`，则：
```xml
<!-- 修改前 -->
<SourceLocation DefaultValue="https://YOUR_USERNAME.github.io/YOUR_REPO/taskpane.html"/>

<!-- 修改后 -->
<SourceLocation DefaultValue="https://zhangsan.github.io/excel-ai-cleaner/taskpane.html"/>
```

修改后提交并推送：
```bash
git add manifest.xml
git commit -m "Update manifest with GitHub Pages URL"
git push
```

#### 4. 启用 GitHub Pages

1. 在 GitHub 仓库页面，点击 **Settings**
2. 左侧菜单找到 **Pages**
3. 在 **Source** 下选择 `Deploy from a branch`
4. **Branch** 选择 `main`，文件夹选择 `/(root)`
5. 点击 **Save**
6. 等待 1-2 分钟，页面顶部会显示你的 GitHub Pages 地址：
   ```
   Your site is live at https://YOUR_USERNAME.github.io/YOUR_REPO/
   ```

#### 5. 在 Excel 中加载插件

现在你可以直接加载部署好的插件：

**Excel 桌面版 (Windows/Mac)**
1. 打开 Excel
2. 点击 **插入** > **获取加载项** 或 **我的加载项**
3. 选择 **上载我的加载项**
4. 输入 manifest.xml 的在线地址：
   ```
   https://YOUR_USERNAME.github.io/YOUR_REPO/manifest.xml
   ```
   或者下载 manifest.xml 文件后选择本地文件上传
5. 插件将在右侧任务窗格打开

**Excel Online**
1. 打开 [Excel Online](https://www.office.com/launch/excel)
2. 打开或创建一个工作簿
3. 点击 **插入** > **加载项**
4. 选择 **管理我的加载项** > **上载我的加载项**
5. 输入在线地址或上传 manifest.xml 文件

---

## 本地开发（可选）

如果你想在本地开发测试：

### 启动本地服务器

```bash
cd /Users/ztzcl/cursor/ai-zhushou
npm install
npm start
```

服务器将在 `http://localhost:3000` 运行。

### 在 Excel 中加载本地版本

#### Excel 桌面版 (Windows/Mac)
1. 打开 Excel
2. 点击 **插入** > **获取加载项** 或 **我的加载项**
3. 选择 **上载我的加载项**
4. 选择本地的 `manifest-local.xml` 文件
5. 插件将在右侧任务窗格打开

#### Excel Online
1. 打开 [Excel Online](https://www.office.com/launch/excel)
2. 打开或创建一个工作簿
3. 点击 **插入** > **加载项**
4. 选择 **管理我的加载项** > **上载我的加载项**
5. 上传 `manifest-local.xml` 文件

---

## 配置 SiliconFlow API

1. 在任务窗格中输入你的 SiliconFlow API Key
2. API 地址默认为：`https://api.siliconflow.cn/v1/chat/completions`
3. 模型名称默认为：`Qwen/Qwen2.5-7B-Instruct`（可更换为其他模型）
4. 点击 **保存配置**

## 使用数据清洗

1. 在 Excel 中选中需要清洗的数据区域
2. 在任务窗格中输入清洗指令（可选）
3. 点击 **清洗选中区域数据**
4. AI 将自动清洗数据并写回选中区域

## 可用模型

SiliconFlow 支持多种模型，常用的有：

| 模型名称 | 说明 |
|---------|------|
| `Qwen/Qwen2.5-7B-Instruct` | 通义千问 7B，速度快 |
| `Qwen/Qwen2.5-72B-Instruct` | 通义千问 72B，能力强 |
| `deepseek-ai/DeepSeek-V2.5` | DeepSeek 大模型 |
| `THUDM/glm-4-9b-chat` | 智谱 GLM-4 9B |

完整模型列表请参考 [SiliconFlow 文档](https://docs.siliconflow.cn/)。

## 示例清洗指令

- `去除所有单元格的前后空格，删除完全空行`
- `统一日期格式为 YYYY-MM-DD`
- `将所有英文转换为大写`
- `删除包含"测试"字样的行`
- `将手机号格式化为 138-xxxx-xxxx 格式`
- `提取邮箱地址，其他内容删除`

## 注意事项

1. **HTTPS 要求**：Office Add-in 必须使用 HTTPS，GitHub Pages 支持 HTTPS
2. **API Key 安全**：配置保存在浏览器 localStorage 中，不会上传到服务器
3. **数据量限制**：单次清洗的数据量不宜过大，建议分批处理
4. **网络要求**：需要能访问 SiliconFlow API

## 文件说明

```
ai-zhushou/
├── manifest.xml      # GitHib Pages 版本的清单（需要修改 URL）
├── taskpane.html     # 任务窗格界面
├── taskpane.js       # 核心逻辑
├── styles.css        # 样式
├── assets/           # 图标资源
│   ├── icon-32.png
│   └── icon-64.png
├── package.json      # npm 配置
└── README.md         # 说明文档
```

## 获取 API Key

1. 访问 [SiliconFlow](https://cloud.siliconflow.cn/)
2. 注册/登录账号
3. 进入控制台创建 API Key

## 常见问题

### Q: 为什么插件无法加载？
A: 检查以下几点：
1. GitHub Pages 是否已启用并部署成功
2. manifest.xml 中的 URL 是否正确替换
3. 浏览器开发者工具 Console 是否有错误信息

### Q: 本地开发和 GitHub Pages 版本有什么区别？
A: GitHub Pages 版本不需要本地服务器，可以在任何电脑直接使用。本地版本用于开发和测试。

### Q: 如何更新插件？
A: 修改代码后推送到 GitHub，GitHub Pages 会自动更新。Excel 中的插件会在下次打开时自动获取最新版本。

## 许可证

MIT
