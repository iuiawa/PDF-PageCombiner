import fitz,sys,os # fitz : 就是PyMuPDF，必须要1.18.9版本
from PIL import Image

print("当前路径：",os.getcwd(),"\nsys数据路径：",sys.argv[0])
dir_path=os.path.dirname(os.path.abspath(sys.argv[0]))
print("当前文件夹路径：",dir_path)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

def pdf2img(pdf_path, zoom_x, zoom_y):
    doc = fitz.open(pdf_path)  # 打开文档
    basename=os.path.splitext(os.path.basename(pdf_path))[0]
    for page in doc:  # 遍历页面
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom_x, zoom_y))  # 将页面渲染为图片
        pix.writePNG(f'{dir_path}/temp_image/{basename}-Page {page.number+1}.png')  # 将图像存储为PNG格式
    doc.close()  # 关闭文档

def get_all_files_paths(directory):  
    # 获取指定目录下的所有文件和目录名  
    files_and_dirs = os.listdir(directory)  
    # 使用列表推导式来构建所有文件的绝对路径  
    # os.path.join(directory, name) 用于组合目录路径和文件名  
    # os.path.isfile(os.path.join(directory, name)) 确保只包含文件，排除目录  
    all_files_paths = [os.path.join(directory, name) for name in files_and_dirs if os.path.isfile(os.path.join(directory, name))]  
    return all_files_paths 

os.system(f"cd {dir_path}")

# 确保临时文件夹存在
try:
    os.mkdir(f"{dir_path}/temp_image")
    os.mkdir(f"{dir_path}/temp_combined")
    os.mkdir(f"{dir_path}/temp_combined_jpg")
except Exception as e:
    pass

if len(sys.argv) < 2:
    print("\n\n\n [ERROR] 必须接受一个文件才可以运行！请检查你的使用方式是否正确")
    os.system("pause")
    sys.exit()

for path_of_pdf in sys.argv[1:]:
    print("开始处理pdf为图片......")

    os.system(f'del /F /Q {dir_path}\\temp_image\\') # 确保临时文件夹下没有多余文件
    os.system(f'del /F /Q {dir_path}\\temp_combined\\') # 确保临时文件夹下没有多余文件
    os.system(f'del /F /Q {dir_path}\\temp_combined_jpg\\') # 确保临时文件夹下没有多余文件

    pdf_doc=fitz.open(path_of_pdf) # 获取目标文件总页数
    all_num=len(pdf_doc) # 注意在只有一个内容时返回的是1不是0
    pdf_doc.close()

    pdf2img(path_of_pdf, zoom_x=4, zoom_y=4) # 执行转换

    print("处理完成，开始处理图片为强制双页/张纸的PDF图片页面")

    # 先判断张数，输出提示然后执行操作

    if all_num % 2 == 1:
        print("单数页数，开始处理")
    else:
        print("双数页数，开始处理。")

    now_page=1

    img_paths = get_all_files_paths(f'{dir_path}/temp_image')
    while now_page<=all_num:
        if now_page+1 <= all_num:
            img1 = Image.open(img_paths[now_page-1])
            img2 = Image.open(img_paths[now_page])  
            
            # 获取图片尺寸  
            width1, height1 = img1.size  
            width2, height2 = img2.size  
            
            # 计算新图片的宽度和高度（取两者中较高的高度）  
            new_width = width1 + width2  
            new_height = max(height1, height2)  
            
            # 创建一个新的图片对象，模式设置为'RGBA'以支持透明背景（如果PNG图片有透明背景）  
            result = Image.new('RGBA', (new_width, new_height))  
            
            # 将图片粘贴到新图片上  
            result.paste(img1, (0, 0, width1, height1))  
            result.paste(img2, (width1, 0, width1 + width2, height2))  
            
            # 旋转九十度。
            result =result.rotate(-90,expand=True)

            # 保存新图片  
            result.save(f'{dir_path}/temp_combined/{now_page}-COMBINED.png')
        if now_page == all_num:
            img1 = Image.open(img_paths[now_page-1])

            # 让下一页纯白
            width, height = img1.size  
            mode = img1.mode
            img2 = Image.new(mode, (width, height), "white")
            
            # 获取图片尺寸  
            width1, height1 = img1.size  
            width2, height2 = img2.size  
            
            # 计算新图片的宽度和高度（取两者中较高的高度）  
            new_width = width1 + width2  
            new_height = max(height1, height2)  
            
            # 创建一个新的图片对象，模式设置为'RGBA'以支持透明背景（如果PNG图片有透明背景）  
            result = Image.new('RGB', (new_width, new_height))  
            
            # 将图片粘贴到新图片上  
            result.paste(img1, (0, 0, width1, height1))  
            result.paste(img2, (width1, 0, width1 + width2, height2))  
            
            # 旋转九十度。
            result =result.rotate(-90,expand=True)

            # 保存新图片  
            result.save(f'{dir_path}/temp_combined/{now_page}-COMBINED.png')
        now_page += 2

    print("开始将图片合成为最终的PDF强制双页文件！")
    basename_VARY_TEMP,kuo_zhan_ming_TEMP_VARY=os.path.splitext(os.path.basename(path_of_pdf))
    out_put_pdf_path = f"{dir_path}/{basename_VARY_TEMP}_output.pdf"

    # 创建一个新的PDF文档对象  
    doc = fitz.Document()  
    target_images_PNG = get_all_files_paths(f'{dir_path}/temp_combined/')

    num_jpg=0
    for img_path in target_images_PNG:
        # 打开 PNG 图片  
        img = Image.open(img_path)  
        # 转换为 JPEG，并设置压缩质量（0-100）  
        img = img.convert('RGB')
        img.save(f"{dir_path}/temp_combined_jpg/{num_jpg}.jpg", 'JPEG', quality=100)
        num_jpg+=1

    target_images = get_all_files_paths(f'{dir_path}/temp_combined_jpg/')
    # 循环添加页面  
    for _ in range(len(target_images)):  
        doc.new_page()  # 添加一个新页面  

    num_now = 0
    for img_target in target_images:
        # 定位到页面
        page = doc[num_now]

        # 插入图片
        rect =  fitz.Rect(0, 0, 595, 842) 
        page.insert_image(rect, filename=target_images[num_now])

        # 保存到输出文档
        num_now+=1

    doc.save(out_put_pdf_path)
    print(f"PDF已输出到{out_put_pdf_path}，将开始处理下一个文件（如果有）")



print("主逻辑运行完成，开始清理临时文件")

# 清理临时文件
try:
    os.system(f'del /F /Q {dir_path}\\temp_image\\') 
    os.system(f'del /F /Q {dir_path}\\temp_combined\\')
    os.system(f'del /F /Q {dir_path}\\temp_combined_jpg\\')
    os.rmdir(f"{dir_path}/temp_image")
    os.rmdir(f"{dir_path}/temp_combined")
    os.rmdir(f"{dir_path}/temp_combined_jpg")
except Exception as e:
    pass


print("\n垃圾已清理 \n\n-----------------------------\n全部文件已处理完毕\n-----------------------------\n")
os.system('pause')
    