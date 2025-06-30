import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

print("\n Welcome to Interactive Data Cleaning & Visualization Tool")

# Step 1: Ask folder path
folder_path = input("\n Enter full folder path where Excel files are located:\n> ").strip()

if not os.path.exists(folder_path):
    print(" Folder path doesn't exist. Please check and try again.")
    exit()

# Step 2: List Excel files in that folder
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]
if not excel_files:
    print(" No Excel files found in the folder.")
    exit()

print("\n Found Excel files:")
for i, file in enumerate(excel_files, 1):
    print(f"{i}. {file}")

file_choices = input("\n Enter file numbers to load (comma-separated, e.g., 1,3,5) or 'all':\n> ")

# Step 3: Load selected files
df_list = []
if file_choices.lower() == 'all':
    selected_files = excel_files
else:
    try:
        indexes = [int(x.strip()) - 1 for x in file_choices.split(',')]
        selected_files = [excel_files[i] for i in indexes]
    except:
        print(" Invalid input. Please enter valid numbers.")
        exit()

for file in selected_files:
    full_path = os.path.join(folder_path, file)
    df = pd.read_excel(full_path)
    df_list.append(df)

combined_df = pd.concat(df_list, ignore_index=True)
print("\n Combined Data (Before Cleaning):")
print(combined_df)

# Step 4: Ask about duplicate removal
remove_dupes = input("\n Do you want to remove duplicate rows? (yes/no):\n> ").strip().lower()
if remove_dupes == 'yes':
    combined_df = combined_df.drop_duplicates()
    print(" Duplicates removed.")

print("\n Final Data:")
print(combined_df)

cleaned_file_name = input("\n Enter a name for the cleaned file (without extension):\n> ").strip()
if not cleaned_file_name:
    cleaned_file_name = "cleaned_data"
    cleaned_file_name += ".xlsx"
    cleaned_file_path = os.path.join(folder_path, cleaned_file_name)
else:
    cleaned_file_name += ".xlsx"
    cleaned_file_path = os.path.join(folder_path, cleaned_file_name)
combined_df.to_excel(cleaned_file_path, index=False)
print(f" Cleaned data saved as: {cleaned_file_path}")

# Step 5: Ask for column names to plot
print("\n Columns Available:", list(combined_df.columns))
x_col = input("  Enter the column name for X-axis:\n> ").strip()
y_col = input("⬆  Enter the column name for Y-axis:\n> ").strip()
if x_col not in combined_df.columns or y_col not in combined_df.columns:
    print(" Invalid column names.")
    exit()

# Step 6: Ask for plot type
print("\n Choose a plot type:")
print("1. Scatter Plot")
print("2. Line Plot")
print("3. Bar Plot")

plot_choice = input("> ").strip()

# Step 7: Plot
plt.figure(figsize=(8, 5))
if plot_choice == '1':
    sns.scatterplot(data=combined_df, x=x_col, y=y_col)
    plt.title("Scatter Plot")
elif plot_choice == '2':
    sns.lineplot(data=combined_df, x=x_col, y=y_col, marker='o')
    plt.title("Line Plot")
elif plot_choice == '3':
    sns.barplot(data=combined_df, x=x_col, y=y_col)
    plt.title("Bar Plot")
else:
    print(" Invalid plot choice.")
    exit()

plt.xlabel(x_col)
plt.ylabel(y_col)
plt.grid(True)
plt.tight_layout()
plt.show()
