from collections import Counter
import xlsxwriter


def generate_werewords_xlsx(filepath_to_words, filepath_to_save_table, per_cell=6):
    with open(filepath_to_words, "r", encoding="utf-8") as fh:
        words = [word.strip() for word in fh.readlines()]

    counter = Counter(words)
    non_unique_words = [word for word, count in counter.items() if count > 1]
    if non_unique_words:
        print("WARNING: There are non-unique words in set: " + ", ".join(non_unique_words))

    with xlsxwriter.Workbook(filepath_to_save_table) as workbook:
        worksheet = workbook.add_worksheet()

        # A little less than standard MTG cards size
        worksheet.set_column('A:C', 5*5.6)
        worksheet.set_default_row(28.1*8.4)

        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 24,
            'rotation': 90
        })

        row = 0
        while words:
            for col in range(3):
                words_for_cell, words = words[:per_cell], words[per_cell:]
                if len(words) < per_cell:
                    print(f"Last batch of words is not complete, dropping last {per_cell - len(words)} words")
                    words = []
                    break
                text_for_cell = "\n".join([f"{idx}. {word}" for idx, word in zip(range(1, per_cell+1), words_for_cell)])
                worksheet.write(row, col, text_for_cell, cell_format)
            row += 1


if __name__ == "__main__":
    generate_werewords_xlsx('words.txt', 'werewords.xlsx')
