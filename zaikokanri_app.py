import streamlit as st
import pandas as pd
from datetime import datetime
import os
import tempfile, shutil
import pythoncom
import win32com.client

# ========= パス設定 =========
BASE_DIR = r"C:\Users\morim\OneDrive\新しいフォルダー"
MATERIAL_MASTER = os.path.join(BASE_DIR, "material_master.xlsx")
STAFF_MASTER = os.path.join(BASE_DIR, "staff_master.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "inventory_log.xlsx")
EXCEL_FILE = os.path.join(BASE_DIR, "原料在庫表.xlsm")

# ========= マスター読込 =========
def read_master(path, keyword):
    try:
        df = pd.read_excel(path)
        col = [c for c in df.columns if keyword in c]
        if not col:
            raise ValueError(f"'{keyword}' を含む列が見つかりません")
        return df[col[0]].dropna().unique().tolist()
    except Exception as e:
        st.error(f"{os.path.basename(path)} 読込エラー: {e}")
        return []

materials = read_master(MATERIAL_MASTER, "原料")
staffs = read_master(STAFF_MASTER, "作業者")

# ========= 履歴ファイル初期化 =========
if not os.path.exists(LOG_FILE):
    df_init = pd.DataFrame(columns=[
        "日時", "棚", "列", "段", "サブ", "材料名",
        "操作", "数量(kg)", "残数(kg)", "作業者", "現在の材料"
    ])
    df_init.to_excel(LOG_FILE, index=False)

# ========= 画面UI =========
st.title("📦 棚在庫管理アプリ")

with st.form("inventory_form"):
    st.subheader("📍 配置情報")
    c1, c2, c3, c4 = st.columns(4)
    with c1: shelf = st.selectbox("棚", [1, 2])
    with c2: row = st.selectbox("列", list(range(1, 20)))
    with c3: level = st.selectbox("段", list(range(1, 5)))
    with c4: sub = st.selectbox("サブ", ["", "1", "2"], index=0)

    # ========= 現在の在庫表示 =========
    try:
        df = pd.read_excel(LOG_FILE)
        df["サブ"] = pd.to_numeric(df["サブ"], errors="coerce").fillna(0).astype(int).replace("nan", "")
        df["棚"] = pd.to_numeric(df["棚"], errors="coerce").fillna(0).astype(int)
        df["列"] = pd.to_numeric(df["列"], errors="coerce").fillna(0).astype(int)
        df["段"] = pd.to_numeric(df["段"], errors="coerce").fillna(0).astype(int)

        df_loc = df[
            (df["棚"] == shelf) &
            (df["列"] == row) &
            (df["段"] == level) &
            (df["サブ"] == sub)
        ]

        if not df_loc.empty:
            last_entry = df_loc.iloc[-1]
            cur_material = last_entry["材料名"]
            cur_stock = float(last_entry["残数(kg)"])
            st.info(f"🧾 現在の在庫：{cur_material}（{cur_stock} kg）")
        else:
            cur_material, cur_stock = None, 0
            st.warning("📭 この棚は空です。")
    except Exception as e:
        st.error(f"在庫情報取得エラー: {e}")
        cur_material, cur_stock = None, 0

    st.subheader("⚙️ 入出庫情報")
    c5, c6, c7, c8 = st.columns(4)
    with c5: operation = st.radio("操作", ["入庫", "出庫"], horizontal=True)
    with c6: material = st.selectbox("材料名", materials)
    with c7: qty = st.number_input("数量 (kg)", min_value=1, max_value=9999, step=1)
    with c8: staff = st.selectbox("作業者名", staffs)

    submitted = st.form_submit_button("登録する")

# ========= 登録処理 =========
if submitted:
    qty_signed = qty if operation == "入庫" else -qty

    try:
        df_old = pd.read_excel(LOG_FILE)

        # --- サブ列を整数化（NaNは0）---
        df_old["サブ"] = pd.to_numeric(df_old["サブ"], errors="coerce").fillna(0).astype(int)

        # --- 入力されたサブ値を整数化（空欄は0として扱う）---
        sub_val = int(sub) if str(sub).isdigit() else 0

        # --- 同じ棚・列・段・サブを抽出 ---
        df_loc = df_old[
            (df_old["棚"] == shelf) &
            (df_old["列"] == row) &
            (df_old["段"] == level) &
            (df_old["サブ"] == sub_val)
        ]

        # --- 現在の残数・材料取得 ---
        cur_stock = df_loc["残数(kg)"].iloc[-1] if not df_loc.empty else 0
        cur_material = df_loc.iloc[-1]["材料名"] if not df_loc.empty else None

        # --- バリデーション ---
        if operation == "出庫" and cur_stock <= 0:
            st.error("❌ 空の棚からは出庫できません。")
        elif cur_material and cur_material != material:
            st.error(f"❌ この棚には別の材料「{cur_material}」が登録されています。")
        else:
            # --- 入出庫反映 ---
            new_stock = cur_stock + qty_signed

            if new_stock < 0:
                st.error("⚠️ 出庫数量が在庫を超えています。")
            else:
                new_entry = pd.DataFrame([{
                    "日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "棚": shelf,
                    "列": row,
                    "段": level,
                    "サブ": sub_val,
                    "材料名": material,
                    "操作": operation,
                    "数量(kg)": qty_signed,
                    "残数(kg)": new_stock,
                    "作業者": staff,
                    "現在の材料": material if new_stock > 0 else ""
                }])

                df_updated = pd.concat([df_old, new_entry], ignore_index=True)

                # --- 保存処理 ---
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    temp_path = tmp.name
                    df_updated.to_excel(temp_path, index=False)
                shutil.move(temp_path, LOG_FILE)

                st.success(f"✅ 登録完了！（残数：{new_stock} kg）")

                # --- Excelマクロ実行 ---
                try:
                    pythoncom.CoInitialize()
                    try:
                        excel = win32com.client.GetActiveObject("Excel.Application")
                    except:
                        excel = win32com.client.Dispatch("Excel.Application")

                    excel.Visible = True
                    excel.Application.Run("原料在庫表.xlsm!履歴追加")
                    pythoncom.CoUninitialize()

                    st.info("📘 原料在庫表へ履歴を転記しました。")

                except Exception as e:
                    st.warning(f"⚠️ マクロ実行エラー: {e}")

    except Exception as e:
        st.error(f"⚠️ 履歴保存に失敗しました: {e}")