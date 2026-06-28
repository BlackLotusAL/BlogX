# BlogX

「拾光集」是一个纯静态的知识社区月度文章汇总站点。页面启动后会自动读取 `data/articles.xlsx`，每个 Excel sheet 代表一个月份。

## 本地启动

在项目目录运行：

```powershell
cd C:\Users\27407\Desktop\BlogX
npm start
```

默认端口是 `5173`，启动后访问：

```text
http://localhost:5173
```

## 局域网访问

`server.mjs` 默认绑定 `0.0.0.0`，局域网内其他设备可以通过本机局域网 IP 访问。

查看本机 IPv4：

```powershell
ipconfig
```

找到当前网络适配器下的 `IPv4 地址`，假设是 `192.168.1.23`，其他设备访问：

```text
http://192.168.1.23:5173
```

注意：如果无法访问，请确认 Windows 防火墙允许 Node.js 或端口 `5173` 入站访问，并确保设备在同一个局域网。

## 指定端口或地址

```powershell
# 指定端口
node server.mjs 8080

# 指定端口和监听地址
node server.mjs 8080 0.0.0.0
```

也可以使用环境变量：

```powershell
$env:PORT=8080
$env:HOST="0.0.0.0"
npm start
```

## 停止服务器

在启动服务器的终端中按：

```text
Ctrl + C
```

如果服务器在后台运行，可以查找并停止占用端口的进程：

```powershell
Get-NetTCPConnection -LocalPort 5173 -State Listen | Select-Object OwningProcess
Stop-Process -Id <OwningProcess> -Force
```

## 部署

这是静态站点，可以直接部署整个目录。需要包含这些文件：

```text
index.html
styles.css
app.js
favicon.svg
assets/
data/
```

如果部署到静态托管平台，入口文件是 `index.html`。如果使用本项目内置 Node 静态服务器，启动命令为：

```powershell
node server.mjs 5173 0.0.0.0
```

## Excel 表头

每个月份 sheet 的第 1 行使用以下表头：

| 表头 | 必填 | 页面位置 |
| --- | --- | --- |
| 标题 | 是 | 文章标题 |
| 摘要 | 是 | 文章摘要 |
| 分类 | 否 | 标题上方的分类徽标 |
| 发布日期 | 否 | 文章元信息 |
| 阅读时间 | 否 | 文章元信息，可填 `12` 或 `12 分钟` |
| 浏览量 | 否 | 文章元信息，可填数字或带单位文本 |
| 文章链接 | 否 | 标题、卡片和右上角外链按钮 |
| 标签 | 否 | 文章底部标签，多个标签用 `;` 分隔 |
| 标签链接 | 否 | 标签点击地址，推荐写成 `标签=网址; 标签=网址` |
| 作者 | 否 | 文章元信息 |
| 排序 | 否 | 数字越小越靠前，留空时按发布日期倒序 |
| 是否显示 | 否 | 填 `否`、`no`、`false`、`0`、`隐藏` 会隐藏该行 |
