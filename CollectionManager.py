import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import pandas as pd
import os
import shutil

def search():    
    def and_search(terms):
        results = rwe.find_target(terms[0])
        for term in terms[1:]:
            next_results = rwe.find_target(term)
            results = results[results.index.isin(next_results.index)]
        return results

    def or_search(terms):
        results = rwe.find_target(terms[0])
        for term in terms[1:]:
            next_results = rwe.find_target(term)
            results = results.append(next_results).drop_duplicates()
        return results
    
    # 搜尋名稱/標籤
    clean_dataview()
    print("搜尋: " + search_entry.get())
    search_target = search_entry.get()

    if "+" in search_target and "|" in search_target:
        messagebox.showerror("搜尋輸入錯誤","不可同時使用「+」和「|」")
        refresh()
    else:    
        data = []

        if "+" in search_target:
            terms = search_target.split('+')
            results = and_search(terms)
        elif "|" in search_target:
            terms = search_target.split('|')
            results = or_search(terms)
        else:
            results = rwe.find_target(search_target)

        for index, row in results.iterrows():
            filtered_row = [item for i, item in enumerate(row) if i != 1]
            data.append(filtered_row)

        # 排序
        data = sorted(data, key=lambda x: x[0])

        # 依據種類選擇
        if combo.get() in ['小說', '漫畫', '動畫']:
            data = [item for item in data if item[1] == combo.get()]

        clean_dataview()
        for item in data:
            tree.insert("", tk.END, values=item)

        search_entry.delete(0, tk.END)


def refresh():
    # 刷新資料
    search_entry.delete(0, tk.END)
    combo.set("請選擇")
    clean_dataview()
    load_all()

def clean_dataview():    
    # 清空表格
    for item in tree.get_children():
        tree.delete(item)

def add_new_item():
    # 輸入資料
    new_window = tk.Toplevel(root)
    new_window.title("新增收藏")
    new_window.geometry("400x300")

    # 下拉式选单
    tk.Label(new_window, text="選擇收藏種類").pack(pady=5)
    
    options = ["小說", "漫畫", "動畫"]
    selected_option = tk.StringVar(new_window)
    selected_option.set("請選擇")  # 設置默認
    
    dropdown = tk.OptionMenu(new_window, selected_option, *options)
    dropdown.pack(pady=5)

    # 路径选择器
    tk.Label(new_window, text="選擇資料夾").pack(pady=5)
    
    folder_path = tk.StringVar()

    def browse_folder():
        folder_selected = filedialog.askdirectory()
        folder_path.set(folder_selected)

    path_entry = tk.Entry(new_window, textvariable=folder_path, width=40)
    path_entry.pack(pady=5)
    
    browse_button = tk.Button(new_window, text="瀏覽...", command=browse_folder)
    browse_button.pack(pady=5)

    # 确认和取消按钮
    button_frame = tk.Frame(new_window)
    button_frame.pack(pady=20)

    def confirm_action():
        if selected_option.get() != "請選擇":
            selected = selected_option.get()
            path = folder_path.get()
            if path:
                rwe.add_collection_to_excel(path,selected)
                new_window.destroy()
                refresh()
            else:
                messagebox.showwarning("警告", "請選擇資料夾")
        else:
            messagebox.showwarning("警告", "請選擇收藏種類")

    confirm_button = tk.Button(button_frame, text="確認", command=confirm_action)
    confirm_button.pack(side=tk.LEFT, padx=10)

    cancel_button = tk.Button(button_frame, text="取消", command=new_window.destroy)
    cancel_button.pack(side=tk.LEFT, padx=10)

def open_file():
    # 開啟檔案
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        name = item['values'][0]  # 取得名稱 (第一列)
        kind = item['values'][1] # 取得種類 (第一列)
        print(f"開啟: {name}")

        matching_rows = rwe.get_file_adress(name,kind)
        matching_rows = matching_rows.values.tolist()
        if len(matching_rows) == 1:
            if (os.path.exists(matching_rows[0][1])):
                os.startfile(matching_rows[0][1])
            else:
                messagebox.showinfo("開啟檔案錯誤","檔案不存在")
        elif len(matching_rows) <= 0:
            messagebox.showinfo("開啟檔案錯誤","檔案不存在")
        else:
            count = 1
            for adress in matching_rows[count-1]:
                if (os.path.exists(matching_rows[count-1][1])):
                    os.startfile(matching_rows[count-1][1])
                else:
                    messagebox.showinfo("開啟檔案錯誤","檔案"+matching_rows[count-1]+"不存在")
                count += 1
    else:
        messagebox.showinfo('開啟檔案', '未選擇任何項目')
        print("未選擇任何項目")

def on_select(event):
    # 選擇檔案種類(小說/漫畫/動畫)
    clean_dataview()
    if combo.get() == '全部':
        data = []

        for index, row in rwe.load_df().iterrows():
            filtered_row = [item for i, item in enumerate(row) if i != 1]
            data.append(filtered_row)

        # 排序
        data = sorted(data,key=lambda x:x[0])

        if search_entry.get() != None:
            data = [item for item in data if search_entry.get() in item[0]]

        for item in data:
            tree.insert("", tk.END, values=item)
    else:
        data = []

        for index, row in rwe.find_kind(combo.get()).iterrows():
            filtered_row = [item for i, item in enumerate(row) if i != 1]
            data.append(filtered_row)

        # 排序
        data = sorted(data,key=lambda x:x[0])

        if search_entry.get() != None:
            data = [item for item in data if search_entry.get() in item[0]]

        for item in data:
            tree.insert("", tk.END, values=item)
    print("選擇: " + combo.get())

def delete_file():
    # 刪除檔案
    print("delete")

    def confirm_delete(action,name,kind):
        # 二次確認
        first_confirm = messagebox.askyesno("確認删除", "确定要刪除嗎？")
        if first_confirm:
            if action == "刪除管理器資料":
                rwe.delete_save(name,kind)
                print("刪除管理器資料")                
                delete_window.destroy()
            elif action == "删除所有":

                matching_rows = rwe.get_file_adress(name,kind)
                matching_rows = matching_rows.values.tolist()
                print(matching_rows)
                if len(matching_rows) == 1:
                    if (os.path.exists(matching_rows[0][1])):                        
                        rwe.delete_save(name,kind)
                        shutil.rmtree(matching_rows[0][1])
                    else:
                        messagebox.showinfo("刪除檔案錯誤","檔案不存在")
                else:
                    messagebox.showinfo("刪除檔案錯誤","檔案不存在或多於1個")
                print("删除所有")
                delete_window.destroy()
        else:
            messagebox.showinfo("取消", "操作已取消")
            delete_window.destroy()

    selected_item = tree.selection()

    if selected_item :
        item = tree.item(selected_item)
        values = item['values']

        delete_window = tk.Toplevel(root)
        delete_window.title("删除")
        delete_window.geometry("300x150")

        del_manager_btn = tk.Button(delete_window, text="刪除管理器資料", command=lambda: confirm_delete("刪除管理器資料",values[0],values[1]))
        del_manager_btn.pack(pady=10)

        del_all_btn = tk.Button(delete_window, text="删除所有", command=lambda: confirm_delete("删除所有",values[0],values[1]))
        del_all_btn.pack(pady=10)

        cancel_btn = tk.Button(delete_window, text="取消", command=delete_window.destroy)
        cancel_btn.pack(pady=10)
    else :
        messagebox.showinfo("刪除","未選擇檔案")


def edit_collection_data():
    # 編輯標籤
    print("edit")

    selected_item = tree.selection()
    if selected_item :
        item = tree.item(selected_item)
        values = item['values']

        adress = rwe.get_file_adress(values[0],values[1])
        adress = adress.values.tolist()

        values.insert(1,adress[0][1])
        
        # 創建新的編輯視窗
        edit_window = tk.Toplevel(root)
        edit_window.title("編輯項目")
        edit_window.geometry("400x300")

        # 創建名稱、路徑、種類和標籤的輸入框
        labels = ["名稱", "路徑", "種類", "標籤1", "標籤2", "標籤3", "標籤4", "標籤5"]
        entries = []

        for i, label in enumerate(labels):
            tk.Label(edit_window, text=label).grid(row=i, column=0, padx=10, pady=5)
            entry = tk.Entry(edit_window, width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, values[i] if i < len(values) else "")  # 將原有數據填入
            entries.append(entry)

        # 確認按鈕
        def save_changes():
            new_values = [entry.get() for entry in entries]
            rwe.edit_save(values,new_values)
            tree.item(selected_item, values=new_values)
            edit_window.destroy()
            refresh()

        save_button = tk.Button(edit_window, text="保存", command=save_changes)
        save_button.grid(row=len(labels), column=0, columnspan=2, pady=10)
    else:
        messagebox.showinfo('編輯檔案', '未選擇任何項目')
        print("未選擇任何項目")

def load_all():
    # 讀取所有存檔
    data = []

    for index, row in rwe.load_df().iterrows():
        filtered_row = [item for i, item in enumerate(row) if i != 1]
        #filtered_row = str(filtered_row)
        data.append(filtered_row)

    # 排序
    data = sorted(data,key=lambda x:x[0])

    for item in data:
        tree.insert("", tk.END, values=item)

class rw_excel():
    def __init__(self):
        if not os.path.exists('save.xlsx'):
            # 如果文件不存在，創建一個新的空白 Excel 文件
            new_df = pd.DataFrame()  # 創建一個空的 DataFrame
            new_df.to_excel('save.xlsx', index=False)  # 保存為 Excel 文件
    def initialization_excel_read(self):
        # 初始化EXCEL讀寫
        self.df = pd.read_excel('save.xlsx', header=None)
    def load_df(self):
        self.initialization_excel_read()
        return self.df
    def find_target(self,target):
        self.initialization_excel_read()
        # 選取第1,4,5,6,7,8項 (列索引從0開始)
        columns_to_search = [0, 3, 4, 5, 6, 7]

        # 在指定的列中搜尋含有target的行
        filtered_rows = self.df[self.df[columns_to_search].apply(lambda row: row.str.contains(target, na=False)).any(axis=1)]

        return filtered_rows
    def find_kind(self,kind):
        self.initialization_excel_read()
        # 在指定的列中搜尋含有target的行
        filtered_rows = self.df[self.df[2].str.contains(kind, na=False)]

        return filtered_rows
    def get_file_adress(self,target,kind):
        self.initialization_excel_read()
        filtered_df = self.df[(self.df[0].str.contains(target, na=False, regex=False)) & (self.df[2].str.contains(kind, na=False, regex=False))]
        
        return filtered_df
    def edit_save(self,target,change):
        self.initialization_excel_read()

        # 遍歷
        for index,row in self.df.iterrows():
            # 找到
            if list(row) == target:
                self.df.loc[index] = change

        self.df.to_excel('save.xlsx', header=None, index=False)
    def add_collection_to_excel(self,head_adress,selected_kind):
        self.initialization_excel_read()
        if not self.df.empty:
            existing_files = self.df[0].tolist()
        else:
            existing_files = []

        # 新增收藏至excel
        folder_names = [name for name in os.listdir(head_adress) if os.path.isdir(os.path.join(head_adress, name))]

        data = []

        for folder_name in folder_names:
            if folder_name not in existing_files:
                folder_path = os.path.join(head_adress, folder_name)
                kind = selected_kind  # 種類
                tag1 = "None"        # 標籤
                tag2 = "None"        
                tag3 = "None"        
                tag4 = "None"        
                tag5 = "None"        

                # 每个资料夹一行数据
                data.append([folder_name, folder_path, kind, tag1, tag2, tag3, tag4, tag5])


        add_df = pd.DataFrame(data)
        add_df = pd.concat([self.df,add_df]).reset_index(drop=True)
        add_df.to_excel("save.xlsx", index=False, header=False)
    def delete_save(self,name,kind):
        print("delete save")
        self.initialization_excel_read()

        condition = (self.df[0] == name) & (self.df[2] == kind)
        indices_to_remove = self.df[condition].index

        self.df = self.df.drop(indices_to_remove)
        self.df.to_excel('save.xlsx', index=False, header=False)
        refresh()


if __name__ == '__main__':

    rwe = rw_excel()

    # 創建主視窗
    root = tk.Tk()
    root.title("Collection Mangager")
    root.geometry("800x400")

    # 創建頂部框架
    top_frame = tk.Frame(root)
    top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

    # 創建開啟和新增按鈕

    add_button = tk.Button(top_frame, text="新增", command=add_new_item)
    add_button.pack(side=tk.RIGHT, padx=5)

    open_button = tk.Button(top_frame, text="開啟", command=open_file)
    open_button.pack(side=tk.RIGHT, padx=5)

    open_button = tk.Button(top_frame, text="編輯", command=edit_collection_data)
    open_button.pack(side=tk.RIGHT, padx=5)

    open_button = tk.Button(top_frame, text="刪除", command=delete_file)
    open_button.pack(side=tk.RIGHT, padx=5)

    # 創建搜尋框、刷新按鈕、下拉式選單
    search_label = tk.Label(top_frame, text="搜尋:")
    search_label.pack(side=tk.LEFT)

    search_entry = tk.Entry(top_frame, width=30)
    search_entry.pack(side=tk.LEFT)

    search_button = tk.Button(top_frame, text="搜尋", command=search)
    search_button.pack(side=tk.LEFT, padx=5)

    refresh_button = tk.Button(top_frame, text="刷新", command=refresh)
    refresh_button.pack(side=tk.LEFT, padx=5)

    combo_label = tk.Label(top_frame, text="選擇選項:")
    combo_label.pack(side=tk.LEFT, padx=(20, 5))

    options = ["全部","小說", "漫畫", "動畫"]
    combo = ttk.Combobox(top_frame, values=options)
    combo.pack(side=tk.LEFT)
    combo.set("請選擇")
    combo.bind("<<ComboboxSelected>>", on_select)

    # 創建顯示用列表（表格）
    tree_frame = tk.Frame(root)
    tree_frame.pack(pady=10, fill=tk.BOTH, expand=True)

    columns = ("名稱", "種類", "標籤1", "標籤2", "標籤3", "標籤4", "標籤5")
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings")

    # 設置每列的標題
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    # 初始化表格
    
    load_all()

    # 添加垂直滾動條
    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=vsb.set)

    tree.pack(fill=tk.BOTH, expand=True)

    # 運行主循環
    root.mainloop()
