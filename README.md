# 🏭 Excel VBA Enterprise Resource Planning (ERP) System

<div align="center">

**A comprehensive Business Management System built with Microsoft Excel and VBA**

[![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![VBA](https://img.shields.io/badge/VBA-0078D4?style=for-the-badge&logo=visual-studio&logoColor=white)](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](LICENSE)

</div>

## 📊 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [System Architecture](#system-architecture)
- [Entity Management](#entity-management)
- [Statistics & Analytics](#statistics--analytics)
- [Technical Highlights](#technical-highlights)
- [Installation](#installation)
- [Usage Guide](#usage-guide)
- [Project Structure](#project-structure)
- [Screenshots](#screenshots)
- [Technologies Used](#technologies-used)
- [Future Enhancements](#future-enhancements)
- [Contributing](#contributing)
- [Authors](#authors)
- [License](#license)

## 🎯 Overview

This project is a **full-featured Enterprise Resource Planning (ERP) system** developed in **Microsoft Excel** using **VBA (Visual Basic for Applications)**. The application is designed to manage the complete lifecycle of an industrial company, from factory administration and human resources to client relationship management and order processing.

The system features:
- 🔐 **Secure Authentication** with role-based access control
- 📁 **Complete CRUD Operations** for all business entities
- 📊 **Advanced Statistical Dashboard** for data-driven decision making
- 📝 **Data Validation** using Regex patterns
- 💾 **Persistent Storage** in Excel worksheets
- 🖥️ **Intuitive GUI** with custom UserForms

## ✨ Key Features

### 🔐 Access Control & Security

- **Secure Login System**
  - Username and password authentication
  - Session management
  - Protected access to sensitive operations

- **Multi-User Support**
  - Different access profiles (Management, Supervision)
  - Role-based permissions
  - User activity tracking

### 📊 Entity Management (CRUD Operations)

Complete Create, Read, Update, and Delete functionality for:

#### 🏭 Factories Module
- Infrastructure management
- Production capacity tracking
- Expense monitoring
- Revenue and billing management
- Factory location and details

#### 👷 Employees Module
- Comprehensive staff database
- Role management (Director, Manager, Engineer, Operator)
- Salary administration
- Factory assignment
- Employee performance tracking

#### 🤝 Clients Module
- Client database with Tax ID (NIF)
- Geographic location tracking
- Feedback and satisfaction ratings
- Purchase history
- Contact information management

#### 📦 Orders Module
- Purchase order tracking
- Cost calculation and VAT management
- Shipping and arrival date tracking
- Profit margin analysis
- Order status management
- Client-order relationships

### 📊 Statistics & Analytics Module

Comprehensive data analysis capabilities:

**Averages**
- Average salaries by role
- Average shipping times
- Average order values by country
- Average client satisfaction ratings

**Extremes (Max/Min)**
- Factory with highest/lowest revenue
- Top performing employees
- Oldest/newest clients
- Largest/smallest orders

**Quantities & Counts**
- Number of employees per factory
- Total orders by region
- Active vs inactive clients
- Order volume trends

## 🏛️ System Architecture

```
Excel VBA ERP System
│
├── Presentation Layer (UserForms)
│   ├── Login Interface
│   ├── Main Dashboard
│   ├── Entity Management Forms
│   └── Statistical Reports
│
├── Business Logic Layer (VBA Modules)
│   ├── Authentication Module
│   ├── CRUD Operations
│   ├── Validation Logic
│   └── Statistical Calculations
│
└── Data Layer (Excel Worksheets)
    ├── Factories Sheet
    ├── Employees Sheet
    ├── Clients Sheet
    ├── Orders Sheet
    └── Users Sheet
```

## 🛠️ Technical Highlights

### Data Validation
- **Regex Pattern Matching**: Validates date formats, tax IDs, and numeric fields
- **Input Sanitization**: Prevents SQL injection-style attacks
- **Business Rule Enforcement**: Ensures data integrity

### User Interface
- **Custom UserForms**: Professional, intuitive design
- **Dynamic Controls**: Context-aware form elements
- **Real-time Feedback**: Instant validation messages

### Search & Filtering
- **Real-time Search**: Instant filtering as you type
- **Multi-criteria Filtering**: Search across multiple fields
- **Sort Capabilities**: Order data by any column

### Performance Optimization
- **Efficient Data Handling**: Optimized worksheet operations
- **Memory Management**: Proper object cleanup
- **Event Optimization**: Disabled screen updating during bulk operations

## 💻 Installation

### Prerequisites

- **Microsoft Excel** (Version 2010 or higher recommended)
- **Windows Operating System** (for full VBA compatibility)
- **Macros Enabled**: Excel security settings must allow macro execution

### Setup Instructions

1. **Download the Application**
   ```bash
   git clone https://github.com/Narva35/Excel-VBA-ERP-System.git
   ```

2. **Locate the Main File**
   - Navigate to the `bin/` folder
   - Find the `.xlsm` file (Excel Macro-Enabled Workbook)

3. **Open in Excel**
   - Double-click the `.xlsm` file
   - If prompted, click **"Enable Content"** or **"Enable Macros"**

4. **First Launch**
   - The system will initialize the database structure
   - Use demo credentials to log in (see below)

## 🚀 Usage Guide

### Logging In

1. Open the Excel file
2. The login form will appear automatically
3. Use any of the demo credentials:

| Username | Password | Role |
|----------|----------|------|
| Paula | paula123 | Management |
| Maria | maria123 | Management |
| Gonçalo | goncalo123 | Supervision |
| Ekumby | ekumby123 | Supervision |

### Navigating the System

**Main Dashboard**
- Access all modules from the central menu
- View system statistics
- Quick access to recent records

**Managing Factories**
1. Click "Factories" from the main menu
2. View existing factories in the list
3. Click "Add New" to create a factory
4. Select a factory and click "Edit" to modify
5. Use "Remove" to delete (with confirmation)

**Managing Employees**
1. Navigate to "Employees" module
2. Add new staff members with all details
3. Assign employees to factories
4. Update salaries and roles
5. Track employee performance

**Managing Clients**
1. Access the "Clients" section
2. Register new clients with Tax ID
3. Update contact information
4. Record feedback and ratings
5. View client purchase history

**Processing Orders**
1. Go to "Orders" module
2. Create new purchase orders
3. Link orders to clients and factories
4. Calculate costs, VAT, and profit margins
5. Track shipping and arrival dates

**Viewing Statistics**
1. Click "Statistics" from main menu
2. Select analysis type (Averages, Extremes, Quantities)
3. Choose specific metric to view
4. Export reports if needed

## 📁 Project Structure

```
Excel-VBA-ERP-System/
│
├── bin/
│   └── Sistema_ERP.xlsm          # Main executable file
│
├── src/
│   ├── Forms/
│   │   ├── frmAcesso.frm            # Login form
│   │   ├── frmPrincipal.frm         # Main dashboard
│   │   ├── frmAdicionarFabrica.frm  # Add factory
│   │   ├── frmAdicionarFuncionario.frm # Add employee
│   │   ├── frmAdicionarCliente.frm  # Add client
│   │   ├── frmAdicionarEncomenda.frm # Add order
│   │   ├── frmFabricas.frm          # Factories list
│   │   ├── frmFuncionarios.frm      # Employees list
│   │   ├── frmClientes.frm          # Clients list
│   │   ├── frmEncomendas.frm        # Orders list
│   │   ├── frmVisualizar.frm        # View details
│   │   └── frmRemover*.frm          # Delete confirmations
│   │
│   └── Modules/
│       ├── Module1.bas              # Main logic
│       ├── Module2.bas              # Helper functions
│       ├── VerificarFormatos.bas    # Validation logic
│       └── extra.bas                # Additional utilities
│
└── README.md
```

## 📸 Screenshots

*Coming soon: Screenshots of the main interface, login screen, and various modules will be added to showcase the system's capabilities.*

## 🛠️ Technologies Used

- **Platform**: Microsoft Excel (2010+)
- **Language**: VBA (Visual Basic for Applications)
- **Database**: Excel Worksheets
- **UI Framework**: UserForms
- **Validation**: Regular Expressions (RegEx)
- **Architecture**: Three-tier (Presentation, Business Logic, Data)

## 🚀 Future Enhancements

Potential improvements for the system:

- [ ] **Cloud Integration**: Connect to cloud databases (Azure, AWS)
- [ ] **Web Version**: Convert to web-based application
- [ ] **Mobile App**: Develop companion mobile application
- [ ] **Advanced Reports**: PDF export and email distribution
- [ ] **API Integration**: Connect with third-party services
- [ ] **Machine Learning**: Predictive analytics for sales
- [ ] **Multi-language Support**: Internationalization
- [ ] **Audit Trail**: Complete change history logging
- [ ] **Barcode/QR Integration**: For inventory management
- [ ] **Email Notifications**: Automated alerts and reminders
- [ ] **Advanced Security**: Two-factor authentication
- [ ] **Dashboard Widgets**: Customizable metrics display

## 🤝 Contributing

Contributions are welcome! If you'd like to improve this project:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 👤 Authors

**Ekumby Travessa**
- GitHub: [@Narva35](https://github.com/Narva35)

*Project developed for the Programming curricular unit*

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⭐ Show Your Support

If you found this project helpful, please give it a ⭐️!

---

<div align="center">

**Built with ❤️ using Excel VBA**

For questions or support, please open an issue on GitHub.

</div>
