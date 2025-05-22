import streamlit as st
import zipfile, io, os, glob
import pandas as pd

st.title("📦 Excel-Zip → Merged Workbook")

# 1️⃣ File uploader
uploaded = st.file_uploader("上传一个 ZIP 文件，里面放你的 .xlsx 文件", type="zip")
if not uploaded:
    st.info("请上传一个 .zip 包")
    st.stop()

# 2️⃣ In-memory unzip
with zipfile.ZipFile(io.BytesIO(uploaded.read())) as z:
    # extract to a temp folder
    temp_dir = "temp_unzip"
    if os.path.exists(temp_dir):
        import shutil; shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    z.extractall(temp_dir)

# 3️⃣ Locate all .xlsx files
xlsx_files = [
    f for f in glob.glob(f"{temp_dir}/**/*.xlsx", recursive=True)
    if not os.path.basename(f).startswith("~$")
]
st.write(f"发现 {len(xlsx_files)} 个 Excel 文件")

if st.button("合并并生成下载链接"):
    # 4️⃣ Read & normalize
    dfs, max_cols = [], 0
    for fpath in xlsx_files:
        df = pd.read_excel(fpath, header=None)
        dfs.append(df)
        max_cols = max(max_cols, df.shape[1])

    for i, df in enumerate(dfs):
        # pad columns
        for _ in range(max_cols - df.shape[1]):
            df[df.shape[1]] = None
        df.columns = [f"col_{j}" for j in range(max_cols)]
        dfs[i] = df

    # 5️⃣ Concat
    final_df = pd.concat(dfs, ignore_index=True)

    # 6️⃣ Export to BytesIO
    towrite = io.BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False)
    towrite.seek(0)

    # 7️⃣ Provide download
    st.download_button(
        label="⬇️ 下载 合并后的 Excel",
        data=towrite,
        file_name="merged_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
