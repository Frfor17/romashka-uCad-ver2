from win32com.client import Dispatch
import os
import time  # для небольшой задержки

print("Создаем цилиндр через Компас API...")

try:
    # Подключение
    kompas = Dispatch("KOMPAS.Application.7")
    kompas.Visible = True
    kompas.HideMessage = True
    
    # Новая деталь
    doc = kompas.Documents.Add(1)  # ksLTT_DocPart
    part = doc.ksPart
    print("✓ Деталь создана")
    
    # Эскиз на плоскости XY (индекс 1)
    sketch = part.NewEntity(2)  # ksLotSketch = 2
    sketch.Init(1)  # ksLPlnXY = 1
    sketch.Build()
    print("✓ Эскиз XY готов")
    
    # Окружность r=1500мм в центре
    circle = sketch.NewEntity(4)  # ksLotCircleAbs = 4  
    circle.Init(0, 0, 1500)  # x=0, y=0, r=1500мм
    circle.Build()
    print("✓ Окружность создана")
    
    # Завершаем эскиз
    sketch.EndEdit()
    print("✓ Эскиз завершен")
    
    # Обновляем
    part.Update()
    doc.Update()
    
    # Сохраняем
    doc.SaveAs(os.path.join(os.getcwd(), "cylinder_test.m3d"))
    print("✅ Цилиндр готов: cylinder_test.m3d")
    print("Откройте файл в Компас — увидите профиль окружности!")
    
except Exception as e:
    print(f"❌ Ошибка: {e}")
