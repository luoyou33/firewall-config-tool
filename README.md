# 防火墙配置翻译工具

一款将Excel规划表转换为防火墙配置脚本的工具。

## 功能特性

- **三种规划表支持**：安全策略规划表、地址对象规划表、服务对象规划表
- **常量与变量定义**：支持自定义前缀和后缀配置
- **行级别配置**：每行可单独定义前缀后缀，灵活适配不同场景
- **默认配置选配**：可选择是否启用策略日志、会话日志、AV/IPS配置文件
- **独立/批量生成**：支持单独生成或同时导入三张表一起生成
- **一键导出**：配置结果一键复制和下载
- **模板下载**：提供空白Excel模板下载，包含示例数据
- **响应式界面**：支持拖拽上传，主背景色RGB #426ACE

## 快速开始

### 方式一：运行EXE文件（推荐）

直接运行 `dist/防火墙配置翻译工具.exe`，程序会自动启动本地服务器并打开浏览器。

### 方式二：开发模式

```bash
python -m http.server 8080
```

然后访问 http://localhost:8080

### 方式三：打包EXE

```bash
pyinstaller --onefile --windowed --add-data "index.html;." --add-data "src/app.js;src" --add-data "src/style.css;src" --name "防火墙配置翻译工具" start_server.py
```

## Excel模板格式

### 安全策略规划表

| 安全策略名称 | 源安全区域 | 源IP地址 | 目的安全区域 | 目的IP地址 | 服务 | 动作 | 规则名称前缀 | 规则名称后缀 | 源区域前缀 | 源区域后缀 | 源地址前缀 | 源地址后缀 | 目的区域前缀 | 目的区域后缀 | 目的地址前缀 | 目的地址后缀 | 服务前缀 | 服务后缀 |
|-------------|-----------|---------|-------------|-----------|------|------|-------------|-------------|-----------|-----------|-----------|-----------|-------------|-------------|-----------|-----------|--------|--------|
| 安全策略01 | 办公zone | IP组01 | 服务器zone | IP组02 | TCP443 | 允许 | | | | | | | | | | | | |

### 地址对象规划表

| IP组名称 | IP地址 | 地址集前缀 | 地址集后缀 |
|---------|--------|-----------|-----------|
| IP组01 | 192.168.1.1,192.168.1.2 | | |

### 服务对象规划表

| 服务名称 | 服务端口 | 服务集前缀 | 服务集后缀 |
|---------|---------|-----------|-----------|
| TCP443 | TCP协议443端口 | | |

## 配置输出示例

```
# === 地址对象配置 ===
ip address-set IP组01 type object
  address 0 192.168.1.1 mask 32
  address 1 192.168.1.2 mask 32
#

# === 服务对象配置 ===
ip service-set TCP443 type object
  service 0 protocol tcp source-port 0 to 65535 destination-port 443
#

# === 安全策略配置 ===
rule name 安全策略01
  source-zone 办公zone
  destination-zone 服务器zone
  source-address address-set IP组01
  destination-address address-set IP组02
  service TCP443
  policy logging
  session logging
  profile av default
  profile ips default
  action permit
#
```

## 项目结构

```
firewall-config-tool/
├── index.html                 # 主页面
├── package.json               # 项目配置
├── README.md                  # 项目说明
├── start_server.py            # Python启动脚本
├── src/
│   ├── app.js                # 核心逻辑
│   └── style.css             # 样式文件
├── dist/
│   └── 防火墙配置翻译工具.exe  # 打包后的可执行文件
├── 安全策略规划表.xlsx        # 参考模板
├── 地址对象规划表.xlsx        # 参考模板
├── 服务规划表.xlsx           # 参考模板
└── 输出的防火墙配置.xlsx      # 参考输出
```

## 使用说明

1. **下载模板**：点击页面顶部的模板下载按钮获取空白模板
2. **填写数据**：在Excel中填写规划数据，可参考示例数据格式
3. **上传文件**：选择需要导入的规划表（支持单个或多个）
4. **配置常量**：根据需要自定义前缀、后缀和默认配置项
5. **生成配置**：点击"生成配置"按钮获取防火墙配置脚本
6. **导出结果**：一键复制或下载配置内容

## 技术栈

- HTML5
- CSS3
- JavaScript (ES6+)
- SheetJS (xlsx) - Excel文件解析
- Python 3.x - 本地服务器和打包

## 部署说明

### GitHub Pages部署

```bash
npm install gh-pages --save-dev
npm run build
npm run deploy
```

### 本地EXE部署

直接将 `dist/防火墙配置翻译工具.exe` 分发给用户，双击即可运行。
