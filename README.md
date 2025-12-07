# 电话表整理小工具

一个用于整理小区/道路地址 + 电话号码的 Node.js 小工具。

## 功能简介

- 上传 `.xlsx` 电话表：
  - A 列：地址（格式：`XXX-YYY-ZZZ`，后面可以有备注）
  - B 列：电话1
  - C 列：电话2（可以为空，但会保留整列）
- 设置：
  - 街道 / 小区名（缓存 R）
  - 选择以「小区开头」或「道路开头」
  - 设置变量 A / B / C（如：弄 / 号 / 室）
- 生成新的 Excel，包含：
  - 生成地址
  - 拆出来的备注
  - 电话1、电话2

## 本地运行

```bash
npm install
npm start
```

## Github Commit

```bash
git add .
git commit -m "更新详情信息"
git push origin main
```

## The End ### Comprehension
