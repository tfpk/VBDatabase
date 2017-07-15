'Code Taken from Somerville IT deparment, Adapted/Fixed by @tfpk
'Note on syntax:
'In comments, backticks (`) indicate that content between them are written as real code.
Imports System.Data.OleDb
Public Class Form1
    'These Dim statements all tell VB to name and make space for these variables, which you instansiate (create) later
    'Note they don't contain `()` after the `As OleDb...`, since we do not yet want to instansiate them
    Dim conn As New OleDbConnection
    Dim data_adapter As OleDbDataAdapter
    Dim cmd_builder As New OleDbCommandBuilder

    'This DataSet will be your local copy of the database. If you make changes to this in your program, it won't affect the database until you update it
    Dim dataset As New DataSet()

    Dim currentIndex As Integer = 0

    'An Enum is a nice way of keeping a list of named constants. They act like variables in a class that are assigned to numbers (Normal = 0, Add = 1 etc.).
    Enum Modes
        Normal
        Add
    End Enum
    'While we don't use this right now, it is good practice to keep track of what "mode" your form is in within your program.
    Dim mode As Modes = Modes.Normal

    Public MASTER_TABLE_NAME As Integer = 0

    '`GET_CONNECTION_STRING()` return the string with all the info to connect to the database. Most of it should be treated as constant, though a DATA_SOURCE must be specified
    'Note that a function is like a sub, but it "returns" a value. You can use it sort of like a variable, except you must have `()` after it to "call" it.
    Private Function GET_CONNECTION_STRING()
        Dim TERMINATOR As Char = "; "

        Dim PROVIDER As String = "Microsoft.Jet.OLEDB.4.0"

        'TODO: Insert the link (i.e. `C:\...\name.mdb`) to a valid MS Access database
        Dim DATA_SOURCE As String = ""

        Dim USERNAME As String = "admin"
        Dim PASSWORD As String = ""

        Dim return_string As String = ""
        return_string &= "Provider=" & PROVIDER & TERMINATOR
        return_string &= "Data Source= " & DATA_SOURCE & TERMINATOR
        return_string &= "User Id=" & USERNAME & TERMINATOR
        return_string &= "Password=" & PASSWORD & TERMINATOR
        Return return_string
    End Function

    '`change_move_activity` ensures buttons don't let you navigate "beyond" the number of rows in the database
    Private Sub change_move_activity()
        'Determine if we are at the first or last row
        Dim is_first As Boolean = (currentIndex = 0)
        Dim is_last As Boolean = (currentIndex = dataset.Tables(MASTER_TABLE_NAME).Rows().Count - 1)

        'Enable the buttons based on whether we are not at the first or last row
        btnFirst.Enabled = Not is_first
        btnPrevious.Enabled = Not is_first
        btnNext.Enabled = Not is_last
        btnLast.Enabled = Not is_last
    End Sub

    '`Mode_Add` changes the form to the UI for adding to the Database
    Private Sub Mode_Add()
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtLocation.Text = ""

        btnFirst.Visible() = False
        btnPrevious.Visible() = False
        btnNext.Visible() = False
        btnLast.Visible() = False
        btnAdd.Visible() = False
        btnUpdate.Visible() = False
        btnDelete.Visible() = False
        btnDone.Visible() = True
    End Sub

    '`Mode_Normal()` resets the form to the UI for browsing/updating the database
    Private Sub Mode_Normal()
        btnFirst.Visible() = True
        btnPrevious.Visible() = True
        btnNext.Visible() = True
        btnLast.Visible() = True
        btnAdd.Visible() = True
        btnUpdate.Visible() = True
        btnDelete.Visible() = True
        btnDone.Visible() = False
        If dataset.Tables(MASTER_TABLE_NAME).Rows.Count Then
            FillFields()
        Else
            MessageBox.Show("You have no fields in your database!" + "\n" + "You must add one first.")
            Mode_Add()
        End If
        change_move_activity()
    End Sub

    '`FillFields` Fills the text fields on the form.
    Private Sub FillFields()
        txtFirstName.Text = dataset.Tables(MASTER_TABLE_NAME).Rows(currentIndex).Item("FirstName").ToString()
        txtLastName.Text = dataset.Tables(MASTER_TABLE_NAME).Rows(currentIndex).Item("LastName").ToString()
        txtLocation.Text = dataset.Tables(MASTER_TABLE_NAME).Rows(currentIndex).Item("Location").ToString()
    End Sub

    '`get_next_id()` searches the database for the highest EmployeeID currently used, and returns that value + 1.
    Private Function get_next_id()
        Dim highest As Integer = 0
        Dim column As DataColumn = dataset.Tables(MASTER_TABLE_NAME).Columns("EmployeeID")

        'Iterate through each row, and increase `highest` to the value of that ID, if it is bigger than `highest`
        For Each row As DataRow In dataset.Tables(MASTER_TABLE_NAME).Rows()
            If highest < row(column) Then
                highest = row(column)
            End If
        Next
        Return highest + 1
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'We declared conn above, but this actually instansiates (creates) the connection object
        'Everything should be left as is, except `Data Source =[path-to-file]` should be the location of your database
        Dim CONNECTION_STRING As String = GET_CONNECTION_STRING()
        conn = New OleDbConnection(CONNECTION_STRING)

        'This query is what will be used to fill the dataset (i.e. local copy of the database). 
        Dim tbl_master_select_query As String = "SELECT * FROM tbl_master"

        'This instansiates the DataAdapter, which is the main object to communicate between database and your code.
        data_adapter = New OleDbDataAdapter(tbl_master_select_query, conn)

        'This clones the contents of the database (retrieved using the `tbl_master_select_query`)
        data_adapter.Fill(dataset)

        'This line creates the object that will translate changes to the dataset, when set to the database, into the required SQL commands (UPDATE, INSERT and DELETE)
        cmd_builder = New OleDbCommandBuilder(data_adapter)

        Mode_Normal()
    End Sub

    'Deal with movement among rows

    Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click
        currentIndex = 0
        FillFields()
        change_move_activity()
    End Sub

    Private Sub btnPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious.Click
        currentIndex = currentIndex - 1
        FillFields()
        change_move_activity()
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        currentIndex = currentIndex + 1
        FillFields()
        change_move_activity()
    End Sub

    Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
        currentIndex = dataset.Tables(MASTER_TABLE_NAME).Rows.Count - 1
        FillFields()
        change_move_activity()
    End Sub

    'Update an entry after the user has modified it
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim dr As DataRow
        dr = dataset.Tables(MASTER_TABLE_NAME).Rows(currentIndex)

        dr.BeginEdit()
        dr("FirstName") = txtFirstName.Text
        dr("LastName") = txtLastName.Text
        dr("Location") = txtLocation.Text
        dr.EndEdit()

        data_adapter.Update(dataset)
        dataset.AcceptChanges()
    End Sub

    'Go to Add mode
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Mode_Add()
    End Sub

    'Delete the current record, then return to normal mode.
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim dr As DataRow
        dr = dataset.Tables(MASTER_TABLE_NAME).Rows(currentIndex)
        dr.Delete()
        data_adapter.Update(dataset)
        dataset.AcceptChanges()
        If currentIndex = dataset.Tables(MASTER_TABLE_NAME).Rows.Count Then
            currentIndex -= 1
        End If
        Mode_Normal()
    End Sub

    'After adding a new row, insert it and save changes.
    Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click
        Dim dr As DataRow
        dr = dataset.Tables(MASTER_TABLE_NAME).NewRow()

        dr("FirstName") = txtFirstName.Text
        dr("LastName") = txtLastName.Text
        dr("Location") = txtLocation.Text
        dr("EmployeeID") = get_next_id()
        dataset.Tables(MASTER_TABLE_NAME).Rows.Add(dr)

        data_adapter.Update(dataset)
        dataset.AcceptChanges()

        currentIndex = currentIndex + 1
        Mode_Normal()
    End Sub
End Class
