# AutoReport
- Based on `Jinja2`, use Python to easily create Word report.
- It is hard to find detailed document about rendering `Jinja2` template using `docxtpl` in python, especially on generating tables with complex structure. I've been creating auto-generated report for a while and I will share some of my good prectices about this task.

## Quick Start
- A sample `template.docx` is provided in `/template`, showing the basics of a `Jinja2` template
- A sample `main.py` code is provided in the root directory, showing the basics of rendering report content on
- run `python main.py` to render the sample report, and the result will be output as `report/report.docx`

In general, a report can be splitted into three parts: `text`, `table`, `image`, you can find common examples in the sample template.

---
To be continued...
