# 招投标合规审查与 AI 比对 SaaS 系统

这是一个可部署到 Streamlit Community Cloud 的双 Word 文档审查应用。用户在浏览器中上传招标文件和投标文件，应用通过 DeepSeek/OpenAI 兼容的 Chat Completions API 生成严格 JSON，再在内存中渲染带风险高亮的 Excel 报告。

> 本项目使用 **DeepSeek API Key**，不需要 OpenAI 账户或 OpenAI API Key。依赖中的 `openai` 只是调用 DeepSeek OpenAI 兼容接口的 Python SDK。密钥由 `st.secrets` 在服务端读取，不会出现在页面输入框中。

## 功能

- 提取 `.docx` 正文、表格、嵌套表格、页眉页脚，并补充提取文本框和脚注文字。
- 保留段落/表格行来源标记，便于报告回查原文。
- 长文档自动执行“分块证据提取 → 去重压缩 → 最终交叉核查”，不静默截断全文。
- 输出“缺陷核查记录”和“预估打分表”两张工作表。
- 对致命/废标、扣分/瑕疵、正常/符合三类风险整行着色。
- Excel 全程使用 `io.BytesIO` 生成，不在服务器磁盘保存上传文件或报告。

> 当前版本不对 Word 内图片执行 OCR。扫描页、证书照片、签章和图片中的文字必须人工复核。

## 项目文件

```text
.
├── app.py
├── requirements.txt
├── README.md
└── .gitignore
```

## 本地运行

推荐使用 Python 3.12：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

macOS / Linux：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

本地运行前，新建 `.streamlit/secrets.toml`：

```toml
DEEPSEEK_API_KEY = "你的新密钥"
# 以下两项可省略，省略时使用应用默认值
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-v4-flash"
```

该文件已被 `.gitignore` 排除，禁止提交到 GitHub。启动后应用会自动加载密钥。默认配置为：

- Base URL：`https://api.deepseek.com`
- 模型：`deepseek-v4-flash`

Base URL 和模型名称在页面中只读显示，避免访问者把服务端密钥发送到恶意接口或切换到高成本模型。管理员可通过 `DEEPSEEK_BASE_URL` 和 `DEEPSEEK_MODEL` Secrets 修改。

## 部署到 Streamlit Community Cloud

1. 在 GitHub 新建仓库，把 `app.py`、`requirements.txt`、`README.md` 和 `.gitignore` 提交到仓库根目录。
   不要上传 API Key、真实招投标文件或生成的 Excel；提交前先运行 `git status` 检查暂存区。
2. 推送代码：

   ```bash
   git init
   git add app.py requirements.txt README.md .gitignore
   git commit -m "feat: add bid compliance review app"
   git branch -M main
   git remote add origin https://github.com/<你的用户名>/<仓库名>.git
   git push -u origin main
   ```

3. 打开 [Streamlit Community Cloud](https://share.streamlit.io/)，选择 **Create app**。
4. 选择 GitHub 仓库、`main` 分支和入口文件 `app.py`。
5. 在 **Advanced settings → Secrets** 中配置以下内容，并选择 Python 3.12：

   ```toml
   DEEPSEEK_API_KEY = "你的新密钥"
   # 可选的服务端锁定配置
   DEEPSEEK_BASE_URL = "https://api.deepseek.com"
   DEEPSEEK_MODEL = "deepseek-v4-flash"
   ```

6. 保存设置并部署。应用启动后会显示“API Key 已从 Cloud Secrets 安全加载”，无需用户手动输入。

密钥只能放在 Streamlit Cloud Secrets 或本地 `.streamlit/secrets.toml` 中，禁止硬编码进 `app.py`、README 或提交到 GitHub。

## 安全与使用边界

- API Key 通过 `st.secrets` 在服务端读取，不提供前端输入框，不写入 Excel、日志或项目文件。
- 若应用对外公开，任何访问者都可能消耗该 DeepSeek 账户额度；请限制应用访问权限并监控用量。
- 上传内容会发送到用户填写的 API 服务地址；部署前应确认服务商条款和组织数据安全制度。
- 应用会拒绝空文件、损坏 DOCX、异常 ZIP 结构及超出安全体积的文件。
- AI 结论可能遗漏或误判，不构成法律意见；正式投标前必须由专业人员对照原件复核。
- 长文档会产生多次模型调用，耗时和费用高于短文档。
- 若提示 `model not found`，请把侧边栏模型名称改为 DeepSeek 官方当前列出的模型 ID。

## 官方参考

- [DeepSeek：首次调用 API](https://api-docs.deepseek.com/zh-cn/)
- [DeepSeek：JSON Output](https://api-docs.deepseek.com/zh-cn/guides/json_mode/)
- [Streamlit：部署到 Community Cloud](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)
- [Streamlit：应用依赖](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies)
