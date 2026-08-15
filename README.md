# 钢结构吊装机械措施费计算工具

钢结构工程吊装机械措施费计算工具，依据《钢结构吊装设备效率基本定额》（2015版 / 2021版），自动查表计算吊装台班、工期及措施费用，并支持导出 Excel 报价表。

## 技术栈

- Streamlit（网页界面）
- pandas（数据汇总）
- openpyxl（Excel 导出）

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

浏览器访问 http://localhost:8501

## Streamlit Community Cloud 部署

1. 将本仓库推送/关联到 GitHub；
2. 打开 [cloud.streamlit.io](https://cloud.streamlit.io) 并用 GitHub 账号登录；
3. 点击 **New app**，选择本仓库；
4. Main file path 填写 `app.py`；
5. 点击 **Deploy**，等待 1-2 分钟即可获得公网链接。

## 功能模块

1. 项目基本信息（工程名称、定额版本、合同工期、用钢量等）
2. 单体结构信息（6 种结构类型，自动关联构件调整系数）
3. 机械措施（设备-构件组合，自动查基准效率并计算台班/天数）
4. 辅助时间配置（塔吊爬升、履带吊安拆等）
5. 费用单价（台班单价、进出场费、辅助机械费）
6. 材料措施
7. 备注事项
8. 计算结果（按设备/单体汇总、工期校核、费用合计）与 Excel 导出

## 离线网页版

仓库还包含纯前端离线版本 `钢结构吊装机械措施费计算工具.html`（配套 `xlsx.full.min.js` 本地库），两个文件放在同一目录下，浏览器直接打开即可使用，无需联网、无需安装。
