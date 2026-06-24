; F12 触发的脚本
F12::
{
    ; 移动鼠标到 (200, 40)
    MouseMove(200, 40)
    
    ; 点击鼠标
    Click()
    
    ; 等待 0.1 秒
    Sleep(100)
    
    ; 移动鼠标到 (200, 210)
    MouseMove(200, 210)
    
    ; 等待 0.1 秒
    Sleep(100)
    
    ; 移动鼠标到 (500, 210)
    MouseMove(500, 210)
    
    ; 点击鼠标
    Click()
}

; 颜色对比函数，允许一定的容差（默认±15）
CompareColor(actualColor, targetColor, tolerance := 15) {
    ; 提取 RGB 分量
    actualR := (actualColor >> 16) & 0xFF
    actualG := (actualColor >> 8) & 0xFF
    actualB := actualColor & 0xFF
    
    targetR := (targetColor >> 16) & 0xFF
    targetG := (targetColor >> 8) & 0xFF
    targetB := targetColor & 0xFF
    
    ; 检查各分量是否在容差范围内
    return (Abs(actualR - targetR) <= tolerance && 
            Abs(actualG - targetG) <= tolerance && 
            Abs(actualB - targetB) <= tolerance)
}

; F11 触发的脚本
F11::
{
    ; 模拟键盘按下 ctrl + n
    Send("^n")
    
    ; 等待 0.1 秒
    Sleep(100)
    
    ; 若点 (190, 105) 颜色为 3693D2
    if (CompareColor(PixelGetColor(190, 105), 0x3693D2)) {
        ; 移动鼠标到 (280, 110)
        MouseMove(280, 110)
        ; 等待 0.1 秒
        Sleep(100)
        ; 点击左键
        Click()
        ; 结束脚本
        return
    }
    
    ; 移动鼠标到 (271, 638)
    MouseMove(271, 638)
    ; 等待 0.1 秒
    Sleep(100)
    ; 若 (271, 638) 颜色为 318ECC
    if (CompareColor(PixelGetColor(271, 638), 0x318ECC)) {
        ; 等待 0.1 秒
        Sleep(100)
        ; 点击左键
        Click()
        ; 等待 0.1 秒
        Sleep(100)
        ; 移动到 145, 12
        MouseMove(145, 12)
        ; 等待 0.1 秒
        Sleep(100)
        Click()
        '移动到 170, 70
        MouseMove(170, 70)
        ; 等待 0.1 秒
        Sleep(100)
        Click()
        '移动到 1370, 30
        MouseMove(1370, 30)
        ; 等待 0.1 秒
        Sleep(100)
        Click()        
    }
}
