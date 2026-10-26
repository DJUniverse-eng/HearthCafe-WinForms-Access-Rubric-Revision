Namespace HearthCafe
    Public Class AppUser
        Public Property Id As Integer
        Public Property Username As String = ""
        Public Property DisplayName As String = ""
        Public Property Role As String = "STAFF"
        Public Property PasswordHash As String = ""
        Public Property PasswordSalt As String = ""
        Public Property IsActive As Boolean = True
    End Class

    Public Class MenuProduct
        Public Property Id As Integer
        Public Property Name As String = ""
        Public Property Category As String = "Uncategorized"
        Public Property Price As Decimal
        Public Property Description As String = ""
        Public Property ImagePath As String = ""
        Public Property IsAvailable As Boolean = True
        Public Property IsArchived As Boolean
    End Class
End Namespace
