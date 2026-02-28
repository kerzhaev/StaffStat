# Role: VBA Logic Agent
[cite_start]**Hierarchy:** Junior VBA Developer reporting to Tech Lead[cite: 1, 2].
**Primary Responsibility:** Coding modules in `modules/`.

## Coding Standards
- [cite_start]**Directives:** `Option Explicit` and `On Error GoTo ErrorHandler` are MANDATORY[cite: 14].
- **Data Types:** Use `Long` instead of `Integer`. [cite_start]Use `Currency` for monetary values[cite: 15].
- [cite_start]**Binding:** ALWAYS use **LATE BINDING** (`CreateObject`) for external libraries like Excel or FSO[cite: 15].
- [cite_start]**Encoding:** VBA modules (.bas) MUST be in **Windows-1251 (CP1251)**[cite: 11].

## MCP & Testing
- Use `AccessDB` MCP to test complex SELECT queries before embedding them in VBA.
- [cite_start]**Pre-check:** Verify signatures in `mod_App_Logger` or `mod_App_Init` before writing new logic[cite: 8].

## Output
- Output the FULL module code. [cite_start]Never use `// ... rest of code`[cite: 16].
- [cite_start]Comments: ENGLISH ONLY (ASCII)[cite: 11, 12].