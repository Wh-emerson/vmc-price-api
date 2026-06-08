# vmc-price-api

VMC 系列报价机器人 API，部署到 Vercel 后供企业微信智能机器人回调使用。

## 部署前需要补齐

1. 把 VMC 价格表放到 `api/VMC系列价格表.xlsx`。
2. 在 Vercel 项目环境变量中配置：
   - `WECOM_TOKEN`
   - `WECOM_ENCODING_AES_KEY`
   - `WECOM_RECEIVE_ID`：智能机器人场景通常可留空
   - `PRICE_WORKBOOK`：可选，默认 `VMC系列价格表.xlsx`

## 当前保留内容

- 企业微信 GET 校验、POST 解密和加密回复流程。
- Vercel 路由 `/api/quote`。
- VF 项目复制来的报价引擎骨架。

## 下一步

拿到 VMC 价格表和型号解析规则后，需要替换报价表 sheet 路由、型号拆分规则和价格调整规则。
