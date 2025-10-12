# Color Table Cells
A simple plugin that lets you color table cells in Obsidian. 
Supports both manual coloring and automatic rules based on text or numbers.

By default, table cell colors only show in **Reading Mode**.
But you can also enable them in **Live Preview Mode** if you want.

<img width="622" height="422" alt="Image" src="https://github.com/user-attachments/assets/6707f947-c07b-479a-b93c-244c5670d5e9" />

### How to Color Table Cells

First, switch to **Reading Mode**.

<img width="400" alt="Image" src="https://github.com/user-attachments/assets/1d50b118-15a2-4796-ac0e-bd1a65bc5c44" />

Then **right-click** the cell you want to color and pick a background or text color.

<img width="400" height="247" alt="Image" src="https://github.com/user-attachments/assets/fd61958a-1830-44d3-b48a-6e0ab90126cd" />
<img width="400" height="360" alt="Image" src="https://github.com/user-attachments/assets/cdca4fba-6bf1-4677-98f5-46b120db5183" />

---

### Setting Rules for Table Coloring

You can set up rules to color cells automatically.
For example:

* If a cell **contains a keyword**, it gets a specific color.
* If a cell **has a number** that’s greater, smaller, or equal to a set value, it will color itself based on your rule.

<img width="600" alt="Image" src="https://github.com/user-attachments/assets/b90ea065-2c0f-4d98-bf0d-212f7834497c" />

<img width="400" alt="Image" src="https://github.com/user-attachments/assets/3a8923c4-af87-4f2c-8505-6bc03a31aa6c" />
<img width="400" alt="Image" src="https://github.com/user-attachments/assets/abf2e061-eed9-4bd0-b441-e05f5c603a57" />

---

### How Single Cell Coloring Works

When you pick a color for a cell, the plugin saves it to a small data file.  
It keeps track of the file name, table, row, and column, so when you reopen the note, it knows exactly which cell to recolor.


#### Example `data.json`

```json
{
  "settings": { ... },
  "cellData": {
    "Notes/Math.md": {
      "table_0": {
        "row_1": {
          "col_2": {
            "bg": "#FF9494",
            "color": "#000000"
          }
        }
      }
    }
  }
}
```

The Data basically tells Obsidian that in `Notes/Math.md`, 
the first table (`table_0`) has a cell at row 1, column 2 with:
- Background: `#FF9494`
- Text color: `#000000`

**In short:**  
Each colored cell is stored by its file, table, row, and column in `data.json`.  
When you open the note, the plugin finds them and restores the colors automatically.

---

### Does it work on Mobile Obsidian?
Yes! The colors show up on mobile too. Both the Rules and the Single Cell Colouring, woohoo!
