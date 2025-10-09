# 🏗️ NSCP 2015 Load Combination Generator for STAAD.Pro

<div align="center">

**Automate Your Structural Analysis Workflow**

*Developed by* **Engr. Lowrence Scott D. Gutierrez**  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?style=flat&logo=linkedin)](https://www.linkedin.com/in/lsdg)

---

**Transform hours of manual load combination work into seconds.**

</div>

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
- Zero manual data entry required

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

### **Step 3: Install the Macro as a Custom Tool**

#### **Method A: Direct Installation**

1. Open **STAAD.Pro**
2. Navigate to **Utilities** → **Tools** → **Customize**
3. Click the **Commands** tab
4. Click **New** to create a new command
5. **Configuration:**
   - **Name:** `NSCP 2015 Load Combo Generator`
   - **Command:** Browse to your `.bas` or `.vba` file
   - **Icon:** Choose a recognizable icon (optional)
   - **Shortcut:** Assign a keyboard shortcut like `Ctrl+Shift+L` (optional)
6. Click **OK** to save

#### **Method B: Excel-Based Execution**

1. Open the included Excel file with the macro
2. Enable macros when prompted
3. Keep STAAD.Pro running with your model open
4. Click the **Generate Load Combinations** button in Excel

---

### **Step 4: Execute the Generator**

1. **Launch the tool** from the Utilities tab or your assigned shortcut
2. **Select design method:**
   - Click **Yes** for **LRFD** (Load and Resistance Factor Design)
   - Click **No** for **ASD** (Allowable Stress Design)
3. **For LRFD:** Enter the `Ev` seismic vertical effect factor when prompted
   - Typical value: `0.2` (per NSCP 2015 Section 208.4.2)
4. **Watch the magic happen** ✨

---

### **Step 5: Verify Generated Combinations**

1. In STAAD.Pro, go to the **Load Cases** panel
2. Review the newly created combination series:
   - **Series 101-199:** LRFD without seismic
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
- **101 Series:** Basic gravity + lateral (no seismic)
- **201 Series:** Gravity + seismic effects (includes Ev)
- **301 Series:** Alternate load paths and special cases

#### **ASD Series**
- **401 Series:** Basic combinations (0.6D, D+L, etc.)
- **501 Series:** Alternate combinations with reduced factors

### **Seismic Effect Factor (Ev)**
Per NSCP 2015, the vertical seismic effect is typically:
```
Ev = 0.2 × SDS × D
```
Where SDS is the design spectral response acceleration.

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
