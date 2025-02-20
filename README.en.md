# 📄 Office to PDF Converter

<div align="center">

![GitHub last commit](https://img.shields.io/github/last-commit/nicolasvar18/-Convertidor-a-PDF)
![GitHub](https://img.shields.io/github/license/nicolasvar18/-Convertidor-a-PDF)
![PowerShell](https://img.shields.io/badge/PowerShell-%235391FE.svg?style=flat&logo=powershell&logoColor=white)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Nicolas%20Vargas-blue?style=flat&logo=linkedin)](https://www.linkedin.com/in/nicolas-vargas-956b79166/)

[🇺🇸 English](./README.en.md) | [🇪🇸 Español](./README.md)

</div>

## 📝 Description

Powerful PowerShell script that automates the conversion of Microsoft Office documents to PDF. It processes Word and Excel files recursively in folders, offering an intuitive graphical interface and real-time progress bars.

## ✨ Features

- 🖥️ **Graphical Interface**: Folder selector and notifications using `System.Windows.Forms`
- 📊 **Multiple Formats**: Supports `.doc`, `.docx`, `.xls`, `.xlsx`, `.xlsm`
- 📈 **Progress Bar**: Real-time visualization of the conversion process
- 🔔 **Notifications**: Informative alerts upon completion
- 🛠️ **Customizable**: Clean and documented code for easy adaptation

## 🚀 Quick Start

### Prerequisites

- ✅ Microsoft Office (Word and Excel) installed
- ✅ PowerShell 5.1 or higher
- ✅ Execution permissions configured

### 🔧 Installation

1. **Clone the repository**
   ```powershell
   git clone https://github.com/nicolasvar18/-Convertidor-a-PDF.git
   cd -Convertidor-a-PDF
   ```

2. **Configure execution permissions**
   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Run the script**
   ```powershell
   .\ConvertirPDF.ps1
   ```

### 📦 Create Executable (Optional)

```powershell
# Install PS2EXE
Install-Module -Name PS2EXE -Scope CurrentUser

# Convert to executable
Invoke-ps2exe .\ConvertirPDF.ps1 .\ConvertirPDF.exe
```

## 📸 Demo

<div align="center">
  <img src="demo.gif" alt="Converter Demo" width="600"/>
</div>

## 🤝 Contributing

Contributions are welcome. For major changes:

1. 🍴 Fork the repository
2. 🔧 Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. 💾 Commit your changes (`git commit -m 'Add: AmazingFeature'`)
4. 📤 Push to the branch (`git push origin feature/AmazingFeature`)
5. 📩 Open a Pull Request

## 📄 License

This project is under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

Nicolás Vargas
- 🌐 [Website](https://nicolasvargas.dev)
- 💼 [LinkedIn](https://www.linkedin.com/in/nicolas-vargas-956b79166/)

## 🙏 Acknowledgments

- Microsoft Office COM Objects Documentation
- PowerShell Community
- All contributors

---

<div align="center">
⭐ If this project helped you, don't forget to give it a star!
</div> 