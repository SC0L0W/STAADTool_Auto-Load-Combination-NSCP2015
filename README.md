# 🏗️ NSCP 2015 Load Combination Generator for STAAD.Pro

<div align="center">

**Automate Your Structural Analysis Workflow**

*Developed by* **Engr. Lowrence Scott D. Gutierrez**  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?style=flat&logo=linkedin)](https://www.linkedin.com/in/lsdg)

---

### 📊 Repository Stats

![GitHub Views](https://komarev.com/ghpvc/?username=SC0L0W&label=Repository%20Views&color=0e75b6&style=flat)  
![GitHub Stars](https://img.shields.io/github/stars/SC0L0W/STAADTool_Auto-Load-Combination-NSCP2015?style=flat&color=yellow)  
![Python Version](https://img.shields.io/badge/Python-3.8%2B-blue?style=flat&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=flat)

---

**Transform hours of manual load combination work into seconds.**

</div>

---

<img width="1677" height="838" alt="image" src="https://github.com/user-attachments/assets/ea804152-51b6-43f0-bdcc-10f83d6ccc95" />

---

## 🧮 Special Notes

- It will only work if the primary loads are exactly the same as in the template — no missing and no extra load cases.

- There are two versions available: one for static load combinations and another for RSA (Response Spectrum Analysis) combinations.


---

## ✨ What Makes This Special

This intelligent VBA macro revolutionizes how structural engineers generate load combinations in STAAD.Pro. Fully compliant with **NSCP 2015 standards**, it automatically creates precise load combinations for both **LRFD** (Strength Design) and **ASD** (Allowable Stress Design) methodologies—eliminating human error and saving valuable engineering time.

Whether you're designing high-rises, bridges, or industrial facilities, this tool seamlessly integrates into your STAAD.Pro environment, reading your load cases and generating code-compliant combinations instantly.

---

## 🚀 Key Features

<table>
<tr>
<td width="50%">

### 🔌 **Seamless Integration**
- Direct connection to STAAD.Pro API
- Automatic load case detection and classification
- Minimal manual data entry required

### 🎯 **Intelligent Classification**
- Recognizes load types from case titles
- Supports multiple load categories
- Handles orthogonal effects automatically

</td>
<td width="50%">

### 📊 **Comprehensive Coverage**
- LRFD Series: 101, 201, 301
- ASD Series: Basic & Alternate
- Seismic effects with user-defined Ev factor
- Wind loads in multiple directions

### 🛡️ **Production-Ready**
- Robust error handling
- Input validation
- Clear user prompts and feedback

</td>
</tr>
</table>

---

## 📋 Supported Load Case Titles

> **Important:** This version recognizes the following standardized load case naming conventions:

| Code | Description | Category |
|------|-------------|----------|
| **DL1** | Primary Dead Load | Dead Load |
| **DL2** | Secondary Dead Load | Dead Load |
| **LL** | Reducible Live Load | Live Load |
| **LL1** | Live Load Alternative | Live Load |
| **LL2** | Non-Reducible Live Load | Live Load |
| **LLR** | Roof Live Load | Live Load |
| **EX** | Seismic Load (X-direction) | Seismic |
| **EZ** | Seismic Load (Z-direction) | Seismic |
| **WX** | Wind Load (X-direction) | Wind |
| **WZ** | Wind Load (Z-direction) | Wind |
| **RSX** | Seismic Load (X-direction) | Seismic | (availlable only for v1)
| **RSZ** | Seismic Load (Z-direction) | Seismic | (availlable only for v1)

*💡 Tip: Use these exact titles in your STAAD.Pro model for automatic detection. Future versions will support custom naming schemes.*

---

## 🔧 Prerequisites

- ✅ STAAD.Pro installed and running
- ✅ Active structural model with defined load cases
- ✅ Basic familiarity with VBA macros
- ✅ Understanding of NSCP 2015 load combination requirements

---

## 📖 Getting Started

### **Step 1: Clone or Download This Repository**

```bash
# Clone via Git
git clone https://github.com/yourusername/nscp-load-combo-generator.git

# Or download as ZIP
# Click the green "Code" button → Download ZIP
```
<img width="1331" height="760" alt="image" src="https://github.com/user-attachments/assets/47df1fbe-24d1-4595-8aab-67668e675d9a" />

Extract the files to a convenient location on your computer.

---

### **Step 2: Prepare Your STAAD.Pro Model**

1. **Open your structural model** in STAAD.Pro
2. **Define all load cases** using the standardized naming convention above
3. **Example structure:**
   ```
   Load Cases:
   1. DL1 (Self-weight)
   2. DL2 (Superimposed dead load)
   3. LL (Office live load)
   4. LLR (Roof live load)
   5. EX (Seismic X)
   6. EZ (Seismic Z)
   7. WX (Wind X)
   8. WZ (Wind Z)
   ```

---
<img width="512" height="291" alt="image" src="https://github.com/user-attachments/assets/28708974-8883-46ca-b90c-f33db5a0ad4f" />

### **Step 3: Install the Macro as a Custom Tool**

#### **Method A: Direct Installation**

1. Open **STAAD.Pro**
2. Navigate to **Utilities** → **Tools** → **Customize**
3. Click the **Commands** tab
4. Click **New** to create a new command
5. **Configuration:**
   - **Name:** `NSCP 2015 Load Combo Generator`
   - **Command:** Browse to your `.bas` or `.vbs` file
   - **Icon:** Choose a recognizable icon (optional)
   - **Shortcut:** Assign a keyboard shortcut like `Ctrl+Shift+L` (optional)
6. Click **OK** to save
<img width="1916" height="1025" alt="image" src="https://github.com/user-attachments/assets/f5865712-166e-4f8c-8cd0-9d67f828c1bf" />
<img width="648" height="296" alt="image" src="https://github.com/user-attachments/assets/fd372286-23c3-45a6-a7ad-cfa09b59b8bf" />

#### **Method B: Excel-Based Execution**

1. Open the included Excel file with the macro
2. Enable macros when prompted
3. Keep STAAD.Pro running with your model open
4. Click the **Generate Load Combinations** button in Excel

---

### **Step 4: Execute the Generator**

1. **Launch the tool** from the Utilities tab or your assigned shortcut
<img width="1229" height="526" alt="image" src="https://github.com/user-attachments/assets/a1196113-f885-45d5-80b9-722469153fcb" />
<img width="1509" height="811" alt="image" src="https://github.com/user-attachments/assets/3a65635a-aacf-4eb6-8f73-3d5c1c261a6d" />
3. **Select design method:**
   - Click **Yes** for **LRFD** (Load and Resistance Factor Design)
   <img width="1677" height="838" alt="image" src="https://github.com/user-attachments/assets/257372fc-8f93-4e9a-bb03-9a37eb5caf29" />
   <img width="1601" height="1029" alt="image" src="https://github.com/user-attachments/assets/39c5c1a2-ae58-4fd8-a8aa-5e9b3054918e" />
   - Click **No** for **ASD** (Allowable Stress Design)
   <img width="1516" height="867" alt="image" src="https://github.com/user-attachments/assets/92d44ee4-3191-4041-87d6-92f095def904" />

4. **For LRFD:** Enter the `Ev` seismic vertical effect factor when prompted
   - Typical value: `0.2` (per NSCP 2015 Section 208.4.2)
5. **Watch the magic happen** ✨


---

### **Step 5: Verify Generated Combinations**

1. In STAAD.Pro, go to the **Load Cases** panel
2. Review the newly created combination series:
   - **Series 101-199:** LRFD 
   - **Series 201-299:** LRFD with seismic
   - **Series 301-399:** LRFD alternate series
   - **Series 401-499:** ASD basic combinations
   - **Series 501-599:** ASD alternate combinations
3. Check factors and load assignments for accuracy

---

## 🎨 Customization Options

Want to extend the functionality? Here's how:

### **Add Custom Load Types**
Modify the classification logic in `ClassifyLoadCase()` function:
```vba
ElseIf InStr(1, lcTitle, "SDL", vbTextCompare) > 0 Then
    ClassifyLoadCase = "SuperimposedDL"
```

### **Create New Combination Series**
Add subroutines following the existing pattern:
```vba
Sub CreateCustomSeries(...)
    ' Your custom combination logic
End Sub
```

### **Integrate Additional Codes**
Adapt the factoring system for other standards (IBC, ASCE 7, etc.)

---

## 🎓 Technical Details

### **Load Combination Series Overview**

#### **LRFD Series**
- **101 Series:** Basic gravity + seismic effects
- **201 Series:** Gravity + seismic effects (includes Orthogonal Effects)
- **301 Series:** Gravity + seismic effects (includes Orthogonal Effects and Ev)

#### **ASD Series**
- **401 Series:** Basic combinations
- **501 Series:** Alternate combinations 

### **Seismic Effect Factor (Ev)**
Per NSCP 2015, the vertical seismic effect is typically:
```
Ev = 0.5 × Ca x I × D
```
Where I is importance factor, D is Dead Load and Ca is Acceleration Coefficient
---

## 🤝 Contributing

Found a bug? Want to add features? Contributions are welcome!

1. Fork this repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📞 Support & Contact

**Engr. Lowrence Scott D. Gutierrez**  
📧 Email: *[Available on LinkedIn]*  
💼 LinkedIn: [Connect with me](https://www.linkedin.com/in/lsdg)

For technical support, please open an issue on this repository.

---

## 📄 License

This project is provided **as-is** for educational and professional use.  
**No warranties expressed or implied.**  

Use this tool at your own discretion and always verify output against code requirements.

---

## 🙏 Acknowledgments

- Built with passion for the structural engineering community
- Compliant with NSCP 2015 (National Structural Code of the Philippines)
- Inspired by the need to eliminate repetitive engineering tasks

---

<div align="center">

### ⭐ Star this repository if it saved you time!

**Made with ❤️ for Structural Engineers**

*Because engineering should be about innovation, not repetition.*

</div>
