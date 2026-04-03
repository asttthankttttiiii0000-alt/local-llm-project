import os
import time
from datetime import datetime
from langchain_community.document_loaders import (
    UnstructuredFileLoader,
    UnstructuredWordDocumentLoader,
    UnstructuredExcelLoader,
    UnstructuredPowerPointLoader,
)
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma
from win32com.client import Dispatch

# ショートカットの実体パスを取得
def resolve_shortcut(path):
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    return shortcut.Targetpath

# 許可された拡張子とローダーのマッピング
loader_mapping = {
    (".txt", ".pdf"): UnstructuredFileLoader,
    (".doc", ".docx"): UnstructuredWordDocumentLoader,
    (".xls", ".xlsx", ".xlsm"): UnstructuredExcelLoader,
    (".ppt", ".pptx"): UnstructuredPowerPointLoader,
    (".py", ".js", ".ts", ".cpp", ".java", ".cs"): UnstructuredFileLoader,
}

start_time = time.time()

# docs フォルダ
docs_folder = "./docs"
all_docs = []

def collect_files_from_folder(folder):
    collected = []
    for root, dirs, files in os.walk(folder):
        print(f"📂 フォルダ: {root}")
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            full_path = os.path.join(root, file)
            if ext == ".lnk":
                try:
                    resolved_path = resolve_shortcut(full_path)
                    if os.path.isdir(resolved_path):
                        # フォルダショートカットの場合 → 再帰
                        collected.extend(collect_files_from_folder(resolved_path))
                    elif os.path.isfile(resolved_path):
                        # ファイルショートカットの場合 → そのまま追加
                        collected.append(resolved_path)
                except Exception as e:
                    print(f"⚠️ ショートカットエラー: {e}")
            else:
                collected.append(full_path)
    return collected

# ファイル収集
files = collect_files_from_folder(docs_folder)

# 不要なファイル除外 & ロード
for file_path in files:
    ext = os.path.splitext(file_path)[1].lower()
    loader_class = None
    for extensions, cls in loader_mapping.items():
        if ext in extensions:
            loader_class = cls
            break

    if loader_class:
        try:
            print(f"📄 読み込み中: {file_path}")
            loader = loader_class(file_path)
            docs = loader.load()
            for doc in docs:
                doc.metadata["source"] = os.path.basename(file_path)
            all_docs.extend(docs)
        except Exception as e:
            print(f"❌ 読み込み失敗: {file_path}: {e}")
    else:
        print(f"⏭ 対応外ファイル: {file_path}")

if not all_docs:
    raise ValueError("❌ 対応するファイルが見つかりませんでした。")

# テキスト分割
splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=200)

split_docs = splitter.split_documents(all_docs)

# ベクトルDB上書き保存
embedding = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
db_path = "./vector_db"
if os.path.exists(db_path):
    import shutil
    shutil.rmtree(db_path)

db = Chroma.from_documents(split_docs, embedding, persist_directory=db_path)
db.persist()

elapsed = time.time() - start_time
print(f"\n✅ ドキュメントのベクトル化と保存が完了しました！⏱ {elapsed:.2f} 秒")

