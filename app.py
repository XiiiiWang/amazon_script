import streamlit as st

# 如果你的函数在另一个模块中，确保将其导入
from run_pro import run_program


# 主应用
def main():
    st.title("Excel Link Extractor")

    # 开始运行程序的按钮
    if st.button("Run Program"):
        new_filename = run_program(callback=st.write)
        st.write(f"New Filename: {new_filename}")

        # 如果你想为用户提供一个下载链接，你可以这样做：
        with open(new_filename, "rb") as f:
            bytes_data = f.read()
            st.download_button(
                label="Download Excel File",
                data=bytes_data,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
