import os
import win32com.client

def convert_word_to_html(input_docx, output_html):
    try:
        # COMオブジェクトの生成
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(os.path.abspath(input_docx))

        # Word文書をHTMLに変換
        doc.SaveAs2(os.path.abspath(output_html), FileFormat=8)  # FileFormat=8 はHTML形式

        print(f"WordファイルをHTMLに変換して {output_html} に保存しました.")
    except Exception as e:
        print(f"変換中にエラーが発生しました: {e}")
    finally:
        # COMオブジェクトの解放
        doc.Close()
        word.Quit()

def main():
    # Wordファイルのパスを指定
    input_docx = "your_word_file.docx"

    # 出力HTMLファイルのパスを指定
    output_html = "output.html"

    # Word文書をHTMLに変換
    convert_word_to_html(input_docx, output_html)

if __name__ == "__main__":
    main()
