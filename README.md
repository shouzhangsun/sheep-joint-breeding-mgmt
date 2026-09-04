# 山羊联合育种数据管理系统（桌面版）

山羊育种数据管理工具.离线加强版记事本.数据可视化

一款面向羊场/育种数据初收集的**离线桌面应用**：基于 Flask + SQLite 内核，用 pywebview 封装为原生 Windows 窗口，无需服务器、断网可用。支持配种繁殖、生长测定、系谱档案、品种管理、数据导入导出与档案卡 PDF 生成。优势在于适用于本交群体，抛弃规模化大场养殖管理软件的繁琐，生产端数据采集人性化。

> 适用对象：种羊场、育种站等羊只（品种可自行添加）性能测定与育种数据采集管理。

---

## 一、功能特性

- **羊只档案**：耳号、品种、性别、出生日期、父母号、来源、状态；支持公/母羊差异化档案卡（HTML 查看 + PDF 下载）。
- **配种繁殖**：母羊 × 公羊配种/产羔记录，自动关联系谱；三代系谱树展示。
- **生长测定**：初生 / 3 月龄 / 6 月龄 / 12 月龄体重与体尺，自动计算日增重、校正体重。
- **品种管理**：下拉选择品种；**任何登录用户录入时均可即时新增品种（下拉旁「＋」）**；Excel 导入时自动把新品种登记进品种库。
- **数据导入**：从产仔记录 Excel 自动生成配种/繁殖/羔羊/初生测定；从生长测定 Excel 导入各阶段体尺。
- **看板统计**：种公/母羊数、测定统计、性别分布图表（Chart.js）。
- **导出**：羊只信息、配繁信息、生长测定均支持 Excel 导出。
- **权限**：管理员（admin）可管理用户与品种改名/删除；普通用户可日常录入与查询。

## 二、技术栈

| 层 | 技术 |
|----|------|
| 后端 | Python 3.13 + Flask + SQLite |
| 前端 | Bootstrap 5 + Chart.js（CDN 本地化） |
| 桌面封装 | pywebview 6.x（原生窗口 + 文件选择对话框） |
| 打包 | PyInstaller（32 位 / win32）+ Inno Setup 6（安装包，内置 x86 WebView2 运行时） |
| 报表 | reportlab（档案卡 PDF） |
| 导入 | openpyxl（Excel 读写） |

构建产物为 **32 位 exe**，自带 x86 WebView2 运行时，可在任意 Windows 7+ 机器直接安装使用。

## 三、目录结构

```
sheep_farm_mgmt/
├── app.py                 # Flask 主程序（所有 API 与页面路由）
├── launcher.py            # 桌面端入口：数据目录/单实例/启动 Flask/打开原生窗口
├── templates/             # 页面模板（录入、列表、档案卡、看板等）
├── static/                # 前端资源（bootstrap、chart、Excel 模板）
├── seed/sheep_farm.db     # 内置空数据库模板（首次启动自动复制，无真实数据）
├── resources/             # 构建用：WebView2 运行时安装包（不入库，需自行放置）
├── installer/             # 生成的安装包（作为 GitHub Releases 附件分发，不入库）
├── build.spec / build.bat # PyInstaller 打包配置与一键构建脚本
├── make_seed.py           # 生成内置空库
├── make_cert_and_sign.ps1 # 自签名证书并签名
├── installer.iss          # Inno Setup 安装脚本
├── Languages/             # Inno Setup 中文语言包
├── requirements.txt       # Python 依赖
└── version_info.txt       # 版本信息（文件属性）
```

## 四、下载与安装

安装包在 GitHub **Releases** 页面以附件形式提供（`Setup_山羊联合育种管理系统.exe`，约 202MB，已内置 x86 WebView2 运行时）。

1. 在仓库 **Releases** 页面下载 `Setup_山羊联合育种管理系统.exe`。
2. **右键以管理员身份运行**安装程序。
3. 首次安装会自动部署 **x86 WebView2 运行时**（若系统缺失），按提示完成即可。
4. 从开始菜单 / 桌面快捷方式启动「山羊联合育种管理系统」。


## 五、数据存储位置

所有业务数据保存在本机用户目录，**不与任何服务器通信**：

```
C:\Users\<用户名>\AppData\Roaming\SheepBreeding\sheep_farm.db
```

- 首次启动自动从内置空库复制创建。
- 备份/迁移：直接复制上述 `sheep_farm.db` 文件即可。
- 卸载程序不会删除该数据，需手动清理。

## 六、快速使用

1. 登录后进入主界面，先从「品种管理」确认/补充品种（或直接在下拉点「＋」即时新增）。
2. 「羊只管理 → 新增」录入基础档案；「配种繁殖」「生长测定」按阶段录入。
3. 「数据导入」可按提供的 Excel 模板批量导入产仔记录与生长测定。
4. 羊只档案卡可查看三代系谱、下载 PDF；看板查看群体统计。

## 七、从源码构建（开发者）

### 环境要求
- Windows + Python 3.13（**32 位 win32** 虚拟环境，用于产出 32 位 exe）
- 依赖：见 `requirements.txt`（flask、openpyxl、reportlab、pdfplumber、pywebview）
- [Inno Setup 6](https://jrsoftware.org/isdl.php)（含中文语言包，已提供 `Languages/ChineseSimplified.isl`）
- `resources/MicrosoftEdgeWebView2Setup.exe`：**x86 独立版**（约 179MB，需自行下载放入，不入库）

### 构建步骤
```bash
# 1. 准备 32 位虚拟环境并装依赖
py -3.13-32 -m venv venv32
venv32\Scripts\pip install -r requirements.txt pyinstaller

# 2. 生成内置空库
venv32\Scripts\python make_seed.py

# 3. PyInstaller 打包（32 位单文件）
venv32\Scripts\pyinstaller build.spec --noconfirm

# 4. 自签名 + 签名 exe
powershell -ExecutionPolicy Bypass -File make_cert_and_sign.ps1

# 5. Inno Setup 编译安装包（需先放好 resources/WebView2 安装包）
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```
产物：`dist/SheepBreeding.exe`、`installer/Setup_山羊联合育种管理系统.exe`。

## 八、常见问题

- **安装时报 WebView2 错误**：确认以管理员身份运行；本安装包内置 x86 WebView2，若系统已有其他架构可能被拦截，按提示重装 x86 运行时即可。
- **白屏/打不开**：多为缺少 x86 WebView2 运行时，或之前装过 64 位冲突；用安装包自带的 WebView2 重装。
- **数据在哪**：见第五节，卸载不删数据，重装后可继续用原 `sheep_farm.db`。
- **想换电脑**：复制 `AppData\Roaming\SheepBreeding\` 整个目录到新机同名位置即可。

## 九、许可证与免责声明

- 本项目目前以**内部科研工具**定位发布。如需对外开源，请自行添加 `LICENSE`（建议 MIT 或 GPL，按单位规定选择）。
- 本系统**完全离线运行**，不上传任何数据；使用者须自行妥善保管业务数据文件。
- 数据准确性由录入方负责，软件提供方不对因数据错误导致的决策损失承担责任。

---