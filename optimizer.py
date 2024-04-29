import pandas as pd
import numpy as np
from tqdm import tqdm  # Import tqdm for the progress bar

class TvisteOptimizer:
    def __init__(self, data_file, column_area, column_program, initial_locked_rows=None):
        self.data_file = data_file
        self.column_area = column_area
        self.column_program = column_program
        self.data = pd.read_excel(data_file)
        self.locked_rows = initial_locked_rows if initial_locked_rows is not None else []

    def _last_on_tviste(self):
        data = self.data
        third_rows = list(range(2, len(data), 3))
        other_rows = list(range(len(data)))
        for i in third_rows:
            other_rows.remove(i)

        list_with_rownumbers = []
        for i in other_rows:
            destination = data.iloc[i, data.columns.get_loc(self.column_area)]
            if destination == "Tvistevägen / Ålidhöjd":
                list_with_rownumbers.append(i)

        for i in third_rows:
            destination = data.iloc[i, data.columns.get_loc(self.column_area)]
            if destination != "Tvistevägen / Ålidhöjd":
                if list_with_rownumbers:
                    swap_idx = list_with_rownumbers.pop(0)
                    # Swap the rows
                    data.iloc[[i, swap_idx]] = data.iloc[[swap_idx, i]].values

        self.data = data

    def lock_tviste_rows(self):
        # Automatically lock rows based on a specific condition
        automatically_locked_rows = [idx for idx in range(len(self.data)) if self.data.iloc[idx][self.column_area] == "Tvistevägen / Ålidhöjd" and (idx + 1) % 3 == 0]
        self.locked_rows += automatically_locked_rows
        self.locked_rows = list(set(self.locked_rows))  # Remove duplicates if any
        return self.locked_rows

    def shuffle_except_locked_rows(self):
        locked_indices = self.locked_rows
        unlocked_indices = [i for i in range(len(self.data)) if i not in locked_indices]

        # Shuffle the unlocked indices
        indices_to_shuffle = unlocked_indices.copy()
        np.random.shuffle(indices_to_shuffle)

        # Create a new DataFrame for the result
        new_data = self.data.copy()
        for original, shuffled in zip(unlocked_indices, indices_to_shuffle):
            new_data.iloc[original] = self.data.iloc[shuffled]

        self.data = new_data

    def calculate_optimization_score(self):
        data = self.data
        score = 0
        for i in range(0, len(data), 3):
            if i + 2 < len(data):
                unique_programs = set(data.iloc[i:i+3][self.column_program])
                unique_count = len(unique_programs)
                if unique_count == 3:
                    score += 1
                elif unique_count == 2:
                    score += 0.5
                elif unique_count == 1:
                    score -= 2
        return score

    def save_with_color_formatting(self, filename):
        # Create a Pandas Excel writer using XlsxWriter as the engine
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        self.data.to_excel(writer, index=False, sheet_name='Sheet1')

        # Access the XlsxWriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Define your color formats
        format1 = workbook.add_format({'bg_color': '#FFC7CE'})
        format2 = workbook.add_format({'bg_color': '#C6EFCE'})

        # Apply formatting based on row numbers
        for i in range(len(self.data)):
            if (i // 3) % 2 == 0:
                worksheet.set_row(i + 1, None, format1)  # +1 because of the header row
            else:
                worksheet.set_row(i + 1, None, format2)

        # Close the Pandas Excel writer and output the Excel file
        writer._save()

    def optimize(self, iterations):
        self._last_on_tviste()
        self.lock_tviste_rows()
        max_score = 0
        best_data = None
        progress_bar = tqdm(range(iterations), desc='Optimizing', unit='iteration')
        for _ in progress_bar:
            self.shuffle_except_locked_rows()
            score = self.calculate_optimization_score()
            progress_bar.set_postfix(score=max_score)
            if score > max_score:
                max_score = score
                best_data = self.data.copy()
        self.data = best_data if best_data is not None else self.data
        final_filename = 'final_optimized.xlsx'
        self.save_with_color_formatting(final_filename)
        return final_filename

"""
# Example usage:
initial_locked = []  # Example of user-specified locked rows
optimizer = TvisteOptimizer('Förfestmarathon 2024.xlsx', 'Vilket område kommer ni vara på?', 'Viket program pluggar ni?', initial_locked)
filename = optimizer.optimize(iterations=100)  # Customize the number of iterations here
print("The final optimized data has been saved in:", filename)
"""