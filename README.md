# 防火墙配置翻译工具

一款将Excel规划表转换为防火墙配置脚本的Web工具。

## 功能特性

- 支持三种规划表的导入：
  - 安全策略规划表
  - 地址对象规划表
  - 服务对象规划表
- 常量与变量定义功能
- 支持独立生成或批量生成配置
- 配置结果一键复制和下载
- 响应式Web界面，支持拖拽上传

## 技术栈

- HTML5
- CSS3
- JavaScript (ES6+)
- SheetJS (xlsx) - Excel文件解析

## 快速开始

### 开发模式

```bash
npm install
npm run dev
```

然后访问 http://localhost:8080

### 构建部署

```bash
npm run build
npm run deploy
```

## Excel模板格式

### 安全策略规划表

| 策略名称 | 动作 | 源区域 | 目的区域 | 源地址 | 目的地址 | 服务 | 描述 |
|---------|------|--------|----------|--------|----------|------|------|
| rule1 | permit | trust | untrust | 192.168.1.0/24 | 10.0.0.0/8 | HTTP | 允许内网访问外网HTTP |

### 地址对象规划表

| 地址名称 | 类型 | IP地址 | 描述 |
|---------|------|--------|------|
| server1 | host | 192.168.1.10 | 服务器1 |

### 服务对象规划表

| 服务名称 | 协议 | 源端口 | 目的端口 | 描述 |
|---------|------|--------|----------|------|
| HTTP | tcp | any | 80 | HTTP服务 |

## 配置输出示例

```
// === 安全策略配置 ===
firewall policy "rule1"
  action permit
  src-zone trust
  dst-zone untrust
  src-address 192.168.1.0/24
  dst-address 10.0.0.0/8
  service HTTP
  log enable
  description "允许内网访问外网HTTP"
exit
```

## 项目结构

```
firewall-config-tool/
├── index.html          # 主页面
├── package.json        # 项目配置
├── README.md          # 项目说明
└── src/
    ├── app.js         # 核心逻辑
    └── style.css      # 样式文件
```

## 部署说明

项目已配置GitHub Pages部署，运行 `npm run deploy` 即可将项目部署到GitHub Pages。
