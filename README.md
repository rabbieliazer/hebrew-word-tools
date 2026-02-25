# Hebrew Word Tools
### **Developed by Tosh Developers**

A high-performance Microsoft Word Add-in designed for Rabbis, educators, and Hebrew language power users. This suite provides professional-grade keyboard layout correction, Atbash ciphers, and a seamless automated update system.

---

## 🚀 Main Features

* **English-Hebrew Converter:** Instantly corrects text typed with the wrong keyboard layout (e.g., converts `vmuo` to `שלום`).
* **Precision Atbash Cipher:** A verified 27-letter Atbash mapping that preserves vowels (Nekudos) and correctly handles final letters (Sofiyot).
* **Smart Selection:** Advanced logic strips trailing paragraph marks and line breaks to prevent document formatting errors.
* **Live Update System:** Users can update to the latest version directly from the Word Ribbon using the "Check Updates" feature.

---

## 📂 File Architecture

| File | Purpose | Location |
| :--- | :--- | :--- |
| **HebrewTools.dotm** | The core Word Add-in containing the Ribbon UI and VBA logic. | Root Directory |
| **version.txt** | A metadata file used by the installer to track the latest release version. | Root Directory |
| **Install_HebrewTools.bat** | A one-click installer that downloads and unblocks the Add-in for Windows users. | Root Directory |
| **auto-version.yml** | A GitHub Action that automatically syncs the version number when the `.dotm` is updated. | `.github/workflows/` |

---

## 🛠 Installation Instructions (For Users)

1.  **Download** the `Install_HebrewTools.bat` file from this repository.
2.  **Close** Microsoft Word completely.
3.  **Run** the `.bat` file. 
    * *Note: The script will automatically download the Add-in to your Word Startup folder and unblock it from Windows Security.*
4.  **Open** Word. You will see a new **Hebrew Tools** tab in your Ribbon.



---

## 👨‍💻 Developer Workflow (Tosh Developers)

This repository is equipped with **Auto-Sync Versioning**. To release an update:

1.  **Modify the Code:** Open `HebrewTools.dotm`, press `ALT + F11`, and update the `Public Const ToolVer` version string.
2.  **Save & Upload:** Save the `.dotm` and upload it to the GitHub main branch.
3.  **Automated Sync:** The GitHub Action will automatically:
    * Unzip the `.dotm` file.
    * Extract the version number from the VBA code.
    * Overwrite `version.txt` with the new version.
4.  **User Update:** All users will be notified of the update the next time they click **Check Updates** in Word.

---

## ⚙️ Technical Requirements

* **Platform:** Microsoft Word for Windows (Office 365, 2021, 2019, 2016).
* **Permissions:** Requires "Enable Macros" to be permitted within Word's Trust Center (handled automatically by the installer).
* **Encoding:** Full Unicode support for Hebrew character range `U+0590` to `U+05FF`.

---

## ⚖️ License & Support
Developed and maintained by **Tosh Developers**. 
For feature requests or bug reports, please submit an issue through the GitHub repository.
