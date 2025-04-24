<div align="center">

# PYwordTester
(Stop filling in manual test reports — start automating them.)
</div>

**PYwordTester** is a Python-based framework designed to automate and evaluate test sequences defined in Word documents. It executes test steps, compares results against specified limits, and generates detailed Word reports. This tool is especially useful for teams seeking consistent testing workflows and documentation during the design validation phase.
## 🎯 Target Audience

This tool is ideal for hardware developers working with hardware prototypes or in low-volume production environments.

## ✨ Key Features

- ✅ Free and open-source — accessible to everyone  
- 📄 Keeps test documentation, execution code, and evaluation results together for maximum readability  
- 🖼️ Supports graphical elements in Word reports (e.g., camera images, oscilloscope screenshots)  
- 📤 Generates ready-to-send reports for individual unit tests  
- 📊 Allows export of data to third-party tools (e.g., databases, Excel) for extended logging and analysis  

---

## ⚙️ Operation

To operate the PYwordTester:

1. Create a `TestContext` and a `TestRunner` object.
2. Link them using a `Controller` element.
3. Load a Word document containing *pre-formatted test tables*.
4. The `TestRunner` extracts Python commands and test limits from the tables.
5. Commands are executed within the `TestContext`.
6. Results are inserted back into the Word table and saved as a report.

---

## 📋 Pre-formatted Test Tables

The `TestRunner` processes only tables that contain **all** of the following tags in at least one of their cells:

| Tag         | Description |
|-------------|-------------|
| `<Title>`   | Name of the test case (value is read from the cell to the right) |
| `<TEST_DATE>` | Filled in automatically with date/time of test execution |
| `<Group>`   | Optional group name for selective execution (e.g., *init*, *program*) |
| `<RUN>`     | Controls test execution: `YES`, `NO`, `ABORT`, or `MUST` |
| `<PYTHON>`  | Python code to execute |
| `<LL>`      | Lower limit (optional) |
| `<UL>`      | Upper limit (optional) |
| `<Actual>`  | Filled with the result of executed code |
| `<Pass/fail>` | Auto-evaluated from comparison between actual result and limits |
| `<KVP>`     | Key-value pair, used for export to external systems |

### 🧩 Formatting Rules

- The runner reads the value **to the right** of `<Title>`, `<TEST_DATE>`, `<Group>`, and `<RUN>`.
- The row containing `<PYTHON>`, `<LL>`, `<UL>`, `<Actual>`, `<Pass/fail>`, and `<KVP>` defines the structure for test steps.
- Any number of test steps can be listed in the rows beneath.
- The area **above** the test rows is flexible and may include documentation.

### 📝 Example Table (in Markdown)

<Title> | This test is a pass | <TEST_DATE> |  | <GROUP> | 1 | <RUN> | RUN
 | (Optional test description) |  |  |  |  |  | 
<KVP> | Action description | <PYTHON> | <LL> | <Actual> | <UL> | Unit | <Pass/fail>
math | Number in range | 3+1 | 2 |  | 4 |  | 
 | Another test | 1+2 | 2 |  | 4 |  | 

 
