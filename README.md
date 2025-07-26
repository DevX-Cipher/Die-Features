# Detect It Easy – COM Shell Extensions

## 📦 Installation Instructions for COM DLLs

To install and register the COM DLLs for **Detect It Easy**, follow these steps:

### 🔧 Registering the DLLs

1. **Open Command Prompt as Administrator.**  
2. **Navigate to the folder** where the DLLs are located.  
3. **Register each DLL** using the following command:
   ```
   regsvr32 YourDllName.dll
   ```
   Replace `YourDllName.dll` with the actual DLL filename you are registering (e.g., `DieProperty.dll`).

### 📁 Deployment Notes

- The DLL you are registering must be placed in the **same directory as Detect It Easy (DIE)**, alongside `die.dll`.
- Make sure any required dependencies (such as Visual C++ Redistributables) are present on the system.
- If you're installing via an MSIX package, the registration process is handled automatically.

🔗 Source code for `die.dll` is available at:  
https://github.com/horsicq/die_library

---

## 🧪 Features

- Adds Explorer Property Sheet or Info Tool Tip

---

## 🐞 Bug Reports & Feedback

If you encounter any bugs or want to suggest improvements, please open an issue or contact me via the project repository:

---
