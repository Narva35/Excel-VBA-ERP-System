# Business Management System (Excel VBA)

This project is a comprehensive Enterprise Resource Planning (ERP) system developed in **Microsoft Excel** using **VBA (Visual Basic for Applications)**. The application manages the entire lifecycle of an industrial company, from factory administration and human resources to client management and order processing, featuring a robust statistical analysis module.

## 🚀 Key Features

### 🔐 Access Control
* **Login System:** Secure entry with username and password authentication.
* **Multi-User Support:** Configured for different access profiles (e.g., Management, Supervision).

### 🗂️ Entity Management (CRUD)
The system allows users to **Add, Edit, View, and Remove** records in the following areas:
* **🏭 Factories:** Management of infrastructure, production capacity, expenses, and billing.
* **👷 Employees:** Database of staff, roles (Director, Manager, Engineer, Operator), salaries, and factory assignment.
* **🤝 Clients:** Client database including Tax ID (NIF), location, feedback ratings, and history.
* **📦 Orders:** Tracking of purchase orders, costs, VAT, shipping/arrival dates, and profit margins.

### 📊 Statistics & Analytics Module
Detailed data analysis to support decision-making, including:
* **Averages:** Average salaries by role, shipping times, averages by country.
* **Extremes (Max/Min):** Identification of outliers (e.g., "Factory with highest revenue", "Oldest Client").
* **Quantities:** Dynamic counts (e.g., number of employees per factory, total orders by region).

### ⚙️ Technical Highlights
* **Data Validation:** Automatic verification of date formats and numeric fields (using Regex).
* **User Interface:** Intuitive navigation through custom UserForms.
* **Search & Filters:** Real-time filtering capabilities within data lists.

## 🛠️ Prerequisites
* Microsoft Excel (Version 2010 or higher recommended).
* Macros must be enabled in Excel security settings.

## 💾 Installation and Usage
1. Download the main file (`.xlsm`) from the `bin/` folder of this repository.
2. Open the file in Excel.
3. Click **"Enable Content"** or **"Enable Macros"** on the yellow warning bar at the top.
4. Use the demo credentials below to log in.

### 🔑 Access Credentials (Demo)
You can use any of the following users to test the application:
* **User:** `Paula`   | **Password:** `paula123`
* **User:** `Maria`   | **Password:** `maria123`
* **User:** `Gonçalo` | **Password:** `goncalo123`
* **User:** `Ekumby`  | **Password:** `ekumby123`

## 👥 Authors
Project developed for the Programming curricular unit.
* **Ekumby Travessa**