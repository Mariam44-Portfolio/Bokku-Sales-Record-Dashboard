# Bokku Sales Record For 2025 And 2026 - Excel Dashboard
## MARIAM ABDULLATEEF PROJECT

## 📋 Project Overview
Built a complete sales record and dashboard for Bokku from raw dataset to interactive Excel dashboard for 2025 and 2026.

## 📂 PART 1 - The Dataset (Excel)
Sheet: `dataset`

Columns:
Product ID | Product | Quantity | Selling Price | Sales Rep | Product Name | Sub Category | Category | Product type | SR First Name | SR Last Name | Sales Rep | Branch

Example products: White slice bread, Perfume, Wheat slice bread, Milk, Iron beans, Chanel, Honey beans, Bio Oil, Indomie, Scotch egg, Nivea Oil, Dove, Ariel, Lattafa, etc.

Categories: Bakery, Provision, Food, Deodorant
Sub Categories: White Bread, Perfume, Beverages, Rice, Lighter, Snacks, Body Spray, Detergent, Cooking oil, etc.
Branches: Ojuelegba, Lawanson, Festac
Sales Reps: Sola Bankole, Ahmed Abubakar, Ago Peters, Bimbo Johnson, Dada Jolasho, Fatai Goriola, John Agassi, Mary Ogbonna, etc.

## 🧠 Excel Formulas Used
```excel
Sales Rep (Full Name) = CONCATENATE([@[SR First Name]],[@[SR Last Name]])

Example: Sola + Bankole = SolaBankole
