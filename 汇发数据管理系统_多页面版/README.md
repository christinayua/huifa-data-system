# 汇发数据管理系统（Streamlit 多页面版）

## 功能

- 销售数据更新
- 服务站数据更新
- 应收账款更新

## 本地运行

```bash
cd 汇发数据管理系统_多页面版
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

浏览器访问终端显示的网址，通常为 `http://localhost:8501`。

## 部署到 Streamlit Community Cloud

1. 新建一个 GitHub 仓库。
2. 把本文件夹中的全部文件上传到仓库根目录。
3. 登录 Streamlit Community Cloud，选择 **Create app**。
4. 选择该 GitHub 仓库和分支。
5. Main file path 填写：`app.py`。
6. 点击 Deploy，等待生成公开网址。

## 添加到 iPhone 主屏幕

1. 使用 Safari 打开部署后的网址。
2. 点击 Safari 底部“分享”按钮。
3. 选择“添加到主屏幕”。
4. 名称可填写“汇发数据”。
5. 点击右上角“添加”。

## 重要说明

`pages/customer_alias.json` 用于主数据系统的人工匹配记忆。在 Streamlit Community Cloud 上，本地文件系统可能随重启或重新部署而恢复，因此长期使用时建议把匹配规则迁移到持久化数据库。
