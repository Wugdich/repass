Attribute VB_Name = "main"
Option Explicit

Sub main()

    With ra_uf.r_type_cb
        .AddItem ("type1")
        .AddItem ("type2")
    End With
    
    With ra_uf.data_type_cb
        .AddItem ("Предварительные")
        .AddItem ("Фактические")
        .AddItem ("Ускоренные")
    End With
    
    ra_uf.Show
    
End Sub
