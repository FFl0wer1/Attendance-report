#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         MKultra69

 Script Function:
	Automatic attendance report in iVMS-4200 AC

#ce ----------------------------------------------------------------------------
; Софтину запускать исключительно от Админа!!! Иначе будут глюки
; Запускаем iVMS
; Пропиши ниже путь до: "iVMS-4200.Framework.C.exe"
Run("HERE")

; Ждем старта 5 сек
Sleep(5000)

; Нажатие на вкладку Attendance Statistics
  MouseClick("left", 105, 600)

; Ждем пол сек, на всякий случай хз 
Sleep(500)

; Жамкаем на вкладку Report
  MouseClick("left", 90, 444)
Sleep(150)

; Жамкаем на кнопошку First/Last Access
  MouseClick("left", 1055, 155)
Sleep(150)

; Далее будут клики с выбором даты, не буду описывать каждый
; Start time
  MouseClick("left", 532, 978)
Sleep(150)
  Send("{ENTER}")
Sleep(150)
  Send("1")
Sleep(150)
  Send("{ENTER}")

; End time
  MouseClick("left", 916, 978)
Sleep(150)
  Send("{ENTER}")
Sleep(150)  
  Send("31")
Sleep(150)

; Выбираем людишек
  MouseClick("left", 638, 936)
Sleep(150)  
  MouseClick("left", 480, 697)
Sleep(150)
  MouseClick("left", 478, 729)
Sleep(150)
  MouseClick("left", 800, 650)

; Осталось подготовить репорт!
; Тык кнопку Report
  MouseClick("left", 1820, 976)
  
Sleep(210)

; Сейвим в pdf потому что xlcx не хочет открываться в приемке(win7)
  MouseClick("left", 52, 67)
  
 Sleep(150) 
 
; Выбираем окно сохранения
  WinActivate("Save PDF File")  
  
 Sleep(150) 
  
; Ждем, пока окно станет активным
  WinWaitActive("Save PDF File") 

 Sleep(150)

; Тыкаем на путь сохранения
  MouseClick("left", 965, 406)
  
Sleep(410)  
  
; Нажимаем Del для удаления пути
  Send("{DEL}")
  
Sleep(250)  

; Втыкаем путь и переходим по нему
; Прописываем путь до папки в которую будет падать репорт
  Send("HERE")
  Send("{ENTER}")

; Тыкаем на имя файла
  MouseClick("left", 1055, 795)

Sleep(250)

; Получаем текущую дату
  $year = @YEAR
  $month = @MON
  $day = @MDAY

; Делаем имя в формате для удобства(ГГГГ-ММ-ДД)
  $date = $year & "-" & $month & "-" & $day & ".pdf"
 
; Вкалачиваем готовое имя
  Send($date)
  
Sleep(500)
  
; Жамкаем финальную кнопошку Save
  MouseClick("left", 1081, 863)

Sleep(1000)
