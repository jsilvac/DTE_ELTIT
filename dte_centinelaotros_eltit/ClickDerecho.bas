Attribute VB_Name = "ClickDerecho"
Option Explicit

' Api ReleaseCapture
Private Declare Function ReleaseCapture Lib "user32" () As Long

' Recibe como parámetro el menu popup y el botón presionado
' ******************************************************
Sub Show_Menu_PopUp(El_menu As Object, Button As Integer)

    If Button = vbRightButton Then

        Dim El_Form As Form
        
        ' Referencia al formulario para poder _
          utilizar el método PopupMenu
        Set El_Form = El_menu.Parent
        
        'Libera el mouse para que no se despliegue el menú estándar
        ReleaseCapture
        
        ' Despliega el menú propio
        El_Form.PopupMenu El_menu
        
        'Elimina la referencia al formulario
        Set El_Form = Nothing

    End If

End Sub



