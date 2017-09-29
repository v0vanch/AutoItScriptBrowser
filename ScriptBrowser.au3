#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         v0vanch

#ce ----------------------------------------------------------------------------

; { прагма директивы компилятору

; разрешение запускать нескомпилированные скрипты
#pragma compile(AutoItExecuteAllowed, True)
; назначение названия продукта
#pragma compile(ProductName, "ScriptBrowser")
; назначение версии продукта
#pragma compile(ProductVersion, 1.0.0.0)
; назначение версии файла
#pragma compile(FileVersion, 1.0.0.0)
; назначение авторских прав
#pragma compile(LegalCopyright, "v0vanch")

; } прагма директивы компилятору

; { блок включения других файлов

#include <File.au3>
#include <Array.au3>

#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <TreeViewConstants.au3>
#include <ColorConstantS.au3>

#include <GuiTreeView.au3>

; } блок включения других файлов

; стандартная директоря для скриптов
Local $defaultScriptsDir = @ScriptDir & "\Scripts"

; если стандартная директория не существует, тогда
If Not FileExists($defaultScriptsDir) Then
   ; создать ее
   DirCreate($defaultScriptsDir & "\")
EndIF

; включение режима обработки событий системных и пользовательского интерфейса
Opt("GUIOnEventMode", 1)

; переменные главного окна
Local $mainWindowTitle = "ScriptBrowser v1.0"
Local $mainWindowWidth = 400
Local $mainWindowHeight = 300

; создание главного окна
Global $hMainWindow = GUICreate($mainWindowTitle, $mainWindowWidth, $mainWindowHeight)
; назначение функции "OnClose" на событие закрытие окна
GUISetOnEvent($GUI_EVENT_CLOSE, "OnClose")

; создание кнопки "Запустить все"
Local $buttonRunAll = GUICtrlCreateButton("Запустить все", $mainWindowWidth - 100, 0, 100, 35)
; назначение функции "OnButtonRunAll" на событие нажатия кнопки
GUICtrlSetOnEvent($buttonRunAll, "OnButtonRunAll")

; создание кнопки "Запустить выбранные"
Local $buttonRunSel = GUICtrlCreateButton("Запустить выбранные", $mainWindowWidth - 100, 35, 100, 35, $BS_MULTILINE)
GUICtrlSetOnEvent($buttonRunSel, "OnButtonRunSel")

; создание надписи с указанием текущей рабочей директории
Local $labelScriptDir = GUICtrlCreateLabel($defaultScriptsDir & " :", 3, 3)

; создание дерева элементов для отображения иерархии скриптов
Local $treeViewScripts = GUICtrlCreateTreeView(0, 20, $mainWindowWidth - 100, $mainWindowHeight - 20)
; установка фонового цвета элемента дерева элементов
GUICtrlSetBkColor($treeViewScripts, 0xF0F0F0)

#cs $treeViewScriptsIDs
[$i] - ControlID - идентификатор элемента управления
[$i][0] - полный путь к файлу
[$i][1] - наименование файла
[$i][2] - тип элемента: папка или скрипт
[$i][3] - массив идентификаторов дочерних элементов
#ce
Local $treeViewScriptsIDs[$treeViewScripts + 1][4]

$treeViewScriptsIDs[$treeViewScripts][2] = "dir"

; рекурсивное заполнение дерева элементов данными из файловой системы
RecursiveScriptTreeFill($treeViewScripts, $defaultScriptsDir)

; если элементов в дереве элементов меньше 21
If UBound($treeViewScriptsIDs) < 21 Then
   ; развернуть дерево элементов
   _GUICtrlTreeView_Expand($treeViewScripts)
EndIf

; установка состояния главного окна
GUISetState(@SW_SHOW, $hMainWindow)

; установка задержки нажатия клавиш
Opt("SendKeyDelay", 100)

; цикл для снижения нагрузки на процессор
While 1
   Sleep(100)
WEnd

Func RecursiveScriptTreeFill($parent, $dir)
   ; получение списка файлов в указанной директории
   Local $fileList = _FileListToArray($dir, "*", 0, True)
   Local $fileListNames = _FileListToArray($dir)
   Local $fileCount = UBound($fileList)

   ; если есть записи о файлах
   If $fileCount > 0 Then

	  ; увеличение размера массива элементов на столько, сколько файлов найдено
	  ReDim $treeViewScriptsIDs[UBound($treeViewScriptsIDs) + $fileList[0]][4]

	  ; объявление массива дочерних элементов
	  Local $siblings[$fileList[0] + 1] = []
	  ; первый элемент содержт информацию о количестве элементов, кроме себя
	  $siblings[0] = $fileList[0]

	  For $i = 1 To $fileList[0]
		 ; создание элемента в дереве элементов
		 Local $elemID = GUICtrlCreateTreeViewItem($fileListNames[$i], $parent)

		 $siblings[$i] = $elemID

		 $treeViewScriptsIDs[$elemID][0] = $fileList[$i]
		 $treeViewScriptsIDs[$elemID][1] = $fileListNames[$i]

		 ; если текущий файл - директория
		 If FileGetAttrib($fileList[$i]) = "D" Then
			; назначение элементу в дереве картинки папки (результат зависит от системы)
			GUICtrlSetImage($elemID, "shell32.dll", 4)
			$treeViewScriptsIDs[$elemID][2] = "dir"

			; вызов этой же функции для текущей директории
			RecursiveScriptTreeFill($elemID, $fileList[$i] & "\")
		 ; если текущий файл - скрипт
		 Else
			$treeViewScriptsIDs[$elemID][2] = "au3"

			; назначение элементу в дереве картинки файла (результат зависит от системы)
			GUICtrlSetImage($elemID, "shell32.dll", 1)
		 EndIf
	  Next

	  $treeViewScriptsIDs[$parent][3] = $siblings
   EndIf
EndFunc

Func RunWaitScript($scriptName)
   ; запуск указанного скрипта с ожиданием завершения выполнения
   RunWait(@ScriptFullPath & ' /AutoIt3ExecuteScript "' & $scriptName & '"')
EndFunc

Func RunScripts($id)
   ; если указанный элемент имеет тип "скрипт"
   If $treeViewScriptsIDs[$id][2] = "au3" Then
	  ; передать полный путь в функцию для его выполнения
	  RunWaitScript($treeViewScriptsIDs[$id][0])
   ; если указанный элемент имеет тип "папка"
   ElseIf $treeViewScriptsIDs[$id][2] = "dir" Then
	  ; для каждого идентификатора дочернего элемента
	  For $entry In $treeViewScriptsIDs[$id][3]
		 ; вызов этой же функции
		 RunScripts($entry)
	  Next
   EndIf
EndFunc

Func OnButtonRunAll()
   ; выполнить все скрипты
   RunScripts($treeViewScripts)
EndFunc

Func OnButtonRunSel()
   ; если элемент не выбран
   If GUICtrlRead($treeViewScripts) = 0 Then
	  MsgBox($MB_ICONWARNING, "", "Элемент не выбран")
   Else
	  ;выполнить скрипт/скрипты из указанной директории
	  RunScripts(GUICtrlRead($treeViewScripts))
   EndIf
EndFunc

Func OnClose()
   ; если текущее окно главное
   If @GUI_WinHandle = $hMainWindow Then
		 ; удалить пользователький интерфейс главного окна
		 GUIDelete(@GUI_WinHandle)
		 ;завершить выполнение прогарммы
		 Exit
   EndIf
EndFunc
