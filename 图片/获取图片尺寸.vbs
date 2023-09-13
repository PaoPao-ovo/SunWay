' /*
'  * @Description: 请填写简介
'  * @Author: LHY
'  * @Date: 2023-09-13 15:18:50
'  * @LastEditors: LHY
'  * @LastEditTime: 2023-09-13 15:49:30
'  */


' /**
'  * @description: 根据传入的图片路径，获取图片的大小
'  * @return { Float,Float } - Img_Width 图片宽 , Img_Height 图片高
'  * @param { String } - ImagePath 文件路径（需要是图片）
'  */
Function GetImageSize(ByVal ImagePath,ByRef Img_Width,ByRef Img_Height)
    
    '图片路径
    Dim PATH_Image
    
    '图片处理对象
    Dim Obj_ImageFile
    
    '图片宽
    Dim Img_Width
    
    '图片高
    Dim Img_Height
    
    '参数初始化
    PATH_Image = ImagePath
    
    Set Obj_ImageFile = CreateObject("WIA.ImageFile")
    
    '加载图片，获取长宽
    Obj_ImageFile.LoadFile PATH_Image
    
    Img_Width = Obj_ImageFile.Width
    
    Img_Height = Obj_ImageFile.Height
    
End Function' GetImageSize




