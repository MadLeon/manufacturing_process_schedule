
### 背景
本 Repo 是一个 monorepo, dir 文件夹中保存的是一个 DIR template 中的所有代码, 你的工作范围停留在本文件夹. 使用 /init 对 claude code 进行初始化

### 细节
- 添加一个模块, 用于更新数据库
- 点击按钮后, 在 part_attachment 中插入新数据, 如失败则写入 Error 日志
- 使用级联插入
  - job_number: H6
  - oe_number: H7
  - po_number: H8
  - part.drawing_number: C7 空格前面的部分

### part 冲突处理
- 通过图纸名调用part时, 可能有多个同名part, 确保调用revision最新的那条
- 在sql中使用降序排列, 并limit最上面的一个
- 应用: Edit Parts 界面; All PO 界面的展开部分

### 数据对齐

part_attachment: 
file_type   "dir"
file_name   "{C7}{空格}@{H6}"
file_path   逻辑参考 DIR Log 文件路径算法
is_active   1


### 参考
[DIR Log 文件路径算法](../dir%20log/mod_CheckSelectedCell.bas)
