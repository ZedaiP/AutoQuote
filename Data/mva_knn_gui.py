import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import numpy as np
from knn_mva_predictor import MVA_KNN_Predictor
import os

class KNNMVAGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("KNN MVA预测器")
        self.root.geometry("800x700")
        
        # 初始化预测器
        self.predictor = MVA_KNN_Predictor(k=3)
        self.model_trained = False
        
        # 创建界面
        self.create_widgets()
        
        # 自动加载数据集
        self.load_data()
    
    def create_widgets(self):
        # 标题
        title_label = ttk.Label(self.root, text="KNN MVA预测器", font=('Arial', 16, 'bold'))
        title_label.pack(pady=10)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 左侧框架 - 输入和训练
        left_frame = ttk.LabelFrame(main_frame, text="模型训练", padding=10)
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 数据集信息
        self.data_info_label = ttk.Label(left_frame, text="正在加载数据...")
        self.data_info_label.pack(pady=5)
        
        # 训练按钮
        train_btn = ttk.Button(left_frame, text="训练模型", command=self.train_model)
        train_btn.pack(pady=10)
        
        # 训练结果
        self.train_result_label = ttk.Label(left_frame, text="")
        self.train_result_label.pack(pady=5)
        
        # 分离线
        ttk.Separator(left_frame, orient='horizontal').pack(fill='x', pady=20)
        
        # 输入区域
        input_label = ttk.Label(left_frame, text="输入特征值进行预测:", font=('Arial', 12, 'bold'))
        input_label.pack(pady=5)
        
        # 创建输入框框架
        self.input_frame = ttk.Frame(left_frame)
        self.input_frame.pack(fill='x', pady=10)
        
        self.input_entries = []
        self.feature_labels = []
        
        # 预测按钮
        predict_btn = ttk.Button(left_frame, text="开始预测", command=self.predict)
        predict_btn.pack(pady=10)
        
        # 右侧框架 - 结果显示
        right_frame = ttk.LabelFrame(main_frame, text="预测结果", padding=10)
        right_frame.pack(side='right', fill='both', expand=True)
        
        # 预测结果文本框
        self.result_text = scrolledtext.ScrolledText(right_frame, height=20, width=50)
        self.result_text.pack(fill='both', expand=True)
        
        # 清空按钮
        clear_btn = ttk.Button(right_frame, text="清空结果", command=self.clear_results)
        clear_btn.pack(pady=5)
    
    def load_data(self):
        """加载数据集"""
        data_path = "Data\\训练.xlsx"
        X, y = self.predictor.load_data(data_path)
        
        if X is None:
            self.data_info_label.config(text="数据加载失败！")
            return
        
        self.X_data = X
        self.y_data = y
        self.data_info_label.config(text=f"数据加载成功: {X.shape[0]}样本, {X.shape[1]}特征")
        
        # 创建输入框
        self.create_input_fields()
    
    def create_input_fields(self):
        """创建输入框"""
        # 清空现有的输入框
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        
        self.input_entries = []
        self.feature_labels = []
        
        if self.predictor.feature_names:
            for i, feature_name in enumerate(self.predictor.feature_names):
                # 特征标签
                label = ttk.Label(self.input_frame, text=f"{feature_name}:")
                label.grid(row=i, column=0, sticky='w', pady=2)
                self.feature_labels.append(label)
                
                # 输入框
                entry = ttk.Entry(self.input_frame, width=15)
                entry.grid(row=i, column=1, sticky='w', pady=2, padx=(5, 0))
                self.input_entries.append(entry)
    
    def train_model(self):
        """训练模型"""
        if not hasattr(self, 'X_data'):
            messagebox.showerror("错误", "数据未加载！")
            return
        
        self.log_result("开始训练模型...")
        
        try:
            results = self.predictor.train(self.X_data, self.y_data)
            
            # 保存模型
            self.predictor.save_model("knn_model.pkl")
            self.model_trained = True
            
            # 显示训练结果
            result_text = f"训练完成！\n"
            result_text += f"R²: {results['r2']:.4f}\n"
            result_text += f"RMSE: {results['rmse']:.4f}\n"
            result_text += f"模型已保存到: knn_model.pkl"
            
            self.train_result_label.config(text=result_text)
            self.log_result(result_text)
            
        except Exception as e:
            error_msg = f"训练失败: {str(e)}"
            self.log_result(error_msg)
            messagebox.showerror("训练错误", error_msg)
    
    def predict(self):
        """执行预测"""
        if not self.model_trained:
            messagebox.showwarning("警告", "请先训练模型！")
            return
        
        try:
            # 获取输入值
            input_values = []
            for entry in self.input_entries:
                value = float(entry.get())
                input_values.append(value)
            
            input_array = np.array(input_values)
            
            # 执行预测
            self.log_result("\n" + "="*50)
            self.log_result("开始预测...")
            
            prediction, neighbor_info = self.predictor.predict_with_neighbors(input_array)
            
            # 显示预测结果
            self.log_result(f"\n预测结果: {prediction:.6f}")
            self.log_result("\n最近的3个邻居:")
            
            for i in range(self.predictor.k):
                neighbor_idx = neighbor_info['indices'][i]
                distance = neighbor_info['distances'][i]
                neighbor_mva = neighbor_info['neighbor_mvas'][i]
                neighbor_features = neighbor_info['neighbor_features'][i]
                
                self.log_result(f"\n邻居 {i+1}:")
                self.log_result(f"  MVA值: {neighbor_mva:.6f}")
                self.log_result(f"  距离: {distance:.6f}")
                self.log_result(f"  特征值: {[f'{val:.3f}' for val in neighbor_features]}")
            
            self.log_result("="*50)
            
        except ValueError as e:
            messagebox.showerror("输入错误", "请确保所有输入都是有效数字！")
        except Exception as e:
            error_msg = f"预测失败: {str(e)}"
            self.log_result(error_msg)
            messagebox.showerror("预测错误", error_msg)
    
    def log_result(self, message):
        """在结果区域显示消息"""
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.see(tk.END)
        self.root.update()
    
    def clear_results(self):
        """清空结果"""
        self.result_text.delete(1.0, tk.END)

def main():
    root = tk.Tk()
    app = KNNMVAGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()