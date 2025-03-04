FEW_SHOT_PROMPT = """
### **Instruction:**
You are given a table in markdown format. Your task is to classify each cell into one of the following categories based on its meaning:

- **HEADER**: A title or label that provides semantic meaning to the table. Headers can appear anywhere.
- **ATTRIBUTE**: An index or row-level descriptor, usually in one or more left-most columns, acting as row headers.
- **DATA**: Actual information corresponding to attributes.
- **NONE**: Empty or missing values in the table.

### **Classification Process:**
Follow these steps before providing the final classification:

1. **Analyze Table Structure**: Identify how the table is organized (e.g., presence of headers, row attributes, data distribution).
2. **Identify Headers**: Determine which cells act as semantic labels (column headers, section headers).
3. **Detect Attributes**: Locate row-level descriptors (usually in the left-most column).
4. **Differentiate Data**: Identify actual values associated with attributes.
5. **Handle Empty Cells**: Mark any missing or empty cells as **NONE**.
6. **Explain Reasoning**: Before outputting the classified table, provide a structured explanation of how the classification was determined.

⚠️ **Important Notes**:
- The input markdown may contain **partial tables, multiple tables, or a single table**.
- Ensure **consistent classification** across different structures.
- **Do NOT modify the number of rows or columns**—the structure must remain identical.
- Base classification **only on the visible content**—do not assume missing context.

---

### **Example 1**
#### **Input Table (Markdown Format)**
{input_example_1}
#### **Output Table (Markdown Format)**
{output_example_1}

---

### **Example 2**
#### **Input Table (Markdown Format)**
{input_example_2}
#### **Output Table (Markdown Format)**
{output_example_2}

---

### **Example 3**
#### **Input Table (Markdown Format)**
{input_example_3}
#### **Output Table (Markdown Format)**
{output_example_3}

---

### **Now, classify the following table in the same way:**

#### **Input Table (Markdown Format)**
{input_table}

#### **Expected Output Format**
1. **Provide a structured reasoning section explaining your classification process.**
2. **Output the classified table in markdown format.**

📝 **Return the final table formatted inside triple backticks (` ```markdown `) to ensure correct rendering.**
"""