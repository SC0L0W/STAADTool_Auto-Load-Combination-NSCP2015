---

# NSCP 2015 Load Combination Generator for STAAD.Pro

**Developed by Engr. Lowrence Scott D. Gutierrez**  
[LinkedIn Profile](https://www.linkedin.com/in/lsdg)

This VBA macro automates the generation of load combinations based on the NSCP 2015 standards for STAAD.Pro. It supports multiple design methods including LRFD (Load and Resistance Factor Design) and ASD (Allowable Stress Design), accommodating various load cases such as dead loads, live loads, wind loads, and seismic loads.

---

## Features

- Connects seamlessly to STAAD.Pro and retrieves load case data.
- Classifies load cases based on their titles (e.g., Dead Loads, Live Loads, Wind, Seismic).
- Supports multiple load combination series:
  - LRFD Series 101, 201, 301 (with and without seismic effects)
  - ASD Series (Basic and Alternate)
- Automates creation and management of load combinations with precise factors.
- Handles orthogonal load effects and seismic effects with user input.
- Includes robust error handling to ensure smooth operation.

---

## Prerequisites

- STAAD.Pro installed and actively running.
- Basic knowledge of VBA macros.
- NSCP 2015 standards for load combinations.

---

## Usage Instructions

### Step 1: Prepare Your STAAD Model
- Load your structural model in STAAD.Pro.
- Define all relevant load cases with appropriate titles (e.g., "DL1", "LL", "EX", "EZ", "WZ", etc.).

### Step 2: Add the Macro as a Custom Tool in STAAD.Pro
1. **Open STAAD.Pro**.
2. Go to the **Utilities** tab.
3. Click on **Tools** → **Customize**.
4. In the **Customize** dialog, select the **Commands** tab.
5. Click **New** to create a new command.
6. Set **Name** to something like "Load Combo Generator".
7. For the **Command**, browse and select your macro file (.bas or .vba script) if applicable, or if your macro is embedded, you may need to copy the macro code here.
8. Alternatively, you can:
   - Save your macro as a `.bas` or `.vba` file.
   - Use STAAD's **Custom Tools** feature to link to your macro.
9. Assign an icon and shortcut key if desired.
10. Click **OK** to add the tool.

### Step 3: Run the Macro via the Custom Tool
- After adding, you'll see your tool in the **Utilities** tab.
- Click your custom tool ("Load Combo Generator") whenever you want to generate load combinations.
- A prompt will appear to select the design method:
  - **Yes** for LRFD (strength design).
  - **No** for ASD (allowable stress design).
- For LRFD, you'll be prompted to enter the `Ev` seismic effect factor.

### Step 4: Generate Load Combinations
- The macro runs and automatically creates the specified load combinations within STAAD.Pro.
- Check the **Load Cases** list to verify the new combinations.

---

## Customization & Extension

- Modify classification logic if your load case titles differ.
- Add new series or load combinations by editing the respective subroutines.
- Integrate additional load effects or standards as needed.

---

## Contact

**Engr. Lowrence Scott D. Gutierrez**  
[LinkedIn Profile](https://www.linkedin.com/in/lsdg)

Feel free to fork, adapt, or extend this macro for your specific project needs!

---

## License

This project is provided **as-is** for educational and professional use. Use at your own discretion.

---

Would you like help with the exact steps to link your macro as a custom tool in STAAD, or do you need a sample script for the custom tool command?
