# 🚀 Ragic ERP Automation & Integration Tools

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/woodylin0920-bit)
[![CI](https://github.com/woodylin0920-bit/erp-tools-ragic/actions/workflows/test.yml/badge.svg)](https://github.com/woodylin0920-bit/erp-tools-ragic/actions/workflows/test.yml)

🌐 [繁體中文說明](README.zh-TW.md)

---

## 📌 Project Overview

A specialized Python toolkit designed to streamline **Ragic Cloud DB** API interactions. This project simplifies the process of automating ERP workflows, enabling efficient data synchronization, automated reporting, and seamless third-party system connectivity.

## ✨ Key Features

* **Advanced API Abstraction**: Encapsulates complex HTTP requests into intuitive Python methods.
* **Workflow Automation**: Minimizes manual data entry and maximizes ERP operational efficiency.
* **Data Pipeline Ready**: Easily integrates into modern data infrastructures, CRM systems, or AI-driven workflows.

## 🛠️ Supported Workflows

| Feature | Description |
|---------|-------------|
| Create Sales Order | Parse customer purchase order Excel files and auto-create Ragic sales orders |
| Create Delivery Order | One-click conversion from sales orders to delivery orders |
| Create Outbound Order | One-click conversion from delivery orders to outbound orders with automatic warehouse data fill-in |
| Export Inventory Report | Pull live warehouse stock from Ragic, auto-convert to PCS, and fill into customer quote template (Excel) |
| Xinzhu Logistics *(coming soon)* | Auto-fill Xinzhu courier system from sales order customer data — eliminate manual copy-paste |
| Agent mode | Natural language interface powered by Claude AI — query inventory, analyze sales trends, get restock recommendations, and export Excel reports |

## 🤖 Agent mode

A conversational AI assistant built into the app. Powered by Claude (Anthropic API).

**▸ Real-time Query**
```
"How much BBB042 is left in TW01?"
"Show me the inventory status across all warehouses."
```

**▸ Intelligence & Analysis**
```
"Which SKUs had the highest sales in Q1?"
"Analyze customer X's order trends and stockout risk."
```

**▸ Supply Chain Suggestion**
```
"How much should I reorder for BBB042 based on the last 3 months?"
"List all SKUs currently below minimum stock levels."
```

**▸ Report Automation**
```
"Export all unfulfilled orders this month to Excel on my desktop."
"Generate a low-stock alert report for TW01."
```

**Setup:** Requires an [Anthropic API key](https://console.anthropic.com). On first launch, the app will prompt you to enter it — stored locally, no config file editing needed.

> Type `reset key` or `重設 key` inside Agent mode to update your API key.

---

## 📋 Templates

Two Excel templates are included for client communication:

| Template | Purpose |
|----------|---------|
| `inventory-template.xlsx` | Full inventory overview — sent to clients to show current Taiwan stock levels; reference images per SKU included |
| `quote-template.xlsx` | Client order form — clients fill in quantities per store (up to 5 stores); includes current stock, order total formula, and reference images per SKU |

### Template Features

- **Reference images** embedded per SKU (sourced from product catalog)
- **Auto-filled inventory** — export function pulls live Ragic stock and fills the 現貨 (stock) column automatically
- **Order total formula** — `=SUM(store columns)` calculates total units ordered per SKU
- **Styled headers** — color-coded sections (product info / order fields / store columns)
- **Output filename** reflects the template used: `quote_TW01_20260405_1430.xlsx`

## 🚀 Quick Start

**Mac:** Double-click `start.command`

**Windows:** Double-click `start.bat`

For first-time setup, refer to the [Chinese documentation](README.zh-TW.md).

---

## 🏭 Industry Experience & Impact

Real-world track record delivering high-reliability, enterprise-grade ERP solutions for fast-growing businesses:

**Case Study: TOYBEBOP INTERNATIONAL CO., LTD.**

- **Scope:** End-to-end Ragic ERP architecture design, data automation integration, and workflow optimization.
- **Convenience Store Expansion:** Optimized operations and data sync logic to support brand entry into **7-ELEVEN and FamilyMart** nationwide networks.
- **Department Store Coverage:** Full integration with **all major department store chains in Taiwan**, including Shin Kong Mitsukoshi, SOGO, Far Eastern, eslite, Mitsui Outlet Park, and LaLaport.
- **Technical Challenge:** Processing high-frequency, complex order data across multiple channel formats with real-time inventory synchronization.

---

## 🤝 Let's Collaborate

I specialize in **ERP Automation** and **Custom System Integrations**. If you're looking to optimize your business efficiency, I offer professional consulting and development services for:

* **Enterprise Architecture Planning & Optimization**
* **Cross-system API Integration** (E-commerce / CRM / Finance)
* **Custom Python Automation Scripts & ERP Solutions**

---

## 📬 Contact

Open to technical discussions and project collaborations:

* **Founder:** Woody Lin
* **Venture:** Whale Spark Global Co., Ltd.
* **Email:** [ceo@whalesparkglobal.com](mailto:ceo@whalesparkglobal.com)
