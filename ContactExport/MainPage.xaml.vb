Imports Microsoft.Phone.UserData
Imports Microsoft.Phone.Tasks

Partial Public Class MainPage
    Inherits PhoneApplicationPage


    ' Constructor
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Save_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim emailAddressChooserTask As EmailAddressChooserTask
        emailAddressChooserTask = New EmailAddressChooserTask()
        AddHandler emailAddressChooserTask.Completed, AddressOf emailAddressChooserTask_Completed
        emailAddressChooserTask.Show()
    End Sub

    Private Sub emailAddressChooserTask_Completed(sender As Object, e As EmailResult)

        If e.TaskResult = TaskResult.OK Then

            Dim json As System.Text.StringBuilder = New System.Text.StringBuilder()
            Dim emailCompose As EmailComposeTask = New EmailComposeTask()

            ' Create a collection of JSON objects
            json.AppendLine("{""contacts"": [ ")

            ' FIXME: JSON doesn't allow trailing commas in arrays

            For Each con As Contact In ContactResultsData.Items
                json.Append("{")

                json.AppendFormat("""displayname"": ""{0}"", ""firstname"": ""{1}"", ""lastname"": ""{2}"", ", con.DisplayName, con.CompleteName.FirstName, con.CompleteName.LastName)

                json.Append("""emails"": [ ")
                For Each email As ContactEmailAddress In con.EmailAddresses
                    json.AppendFormat("""{0}"": ""{1}"",", email.Kind, email.EmailAddress)
                Next
                json.Remove(json.Length - 1, 1)
                json.Append("], ")

                json.Append("""numbers"": [ ")
                For Each phone As ContactPhoneNumber In con.PhoneNumbers
                    json.AppendFormat("""{0}"": ""{1}"",", phone.Kind, phone.PhoneNumber)
                Next
                json.Remove(json.Length - 1, 1)
                json.Append("], ")

                json.Append("""addresses"": [ ")
                For Each address As ContactAddress In con.Addresses
                    json.AppendFormat("""{0}"": {{""addressline1"": ""{1}"", ""city"": ""{2}"", ""countryregion"": ""{3}"", ""postalcode"": ""{4}""}},", address.Kind, address.PhysicalAddress.AddressLine1, address.PhysicalAddress.City, address.PhysicalAddress.CountryRegion, address.PhysicalAddress.PostalCode)
                Next
                json.Remove(json.Length - 1, 1)
                json.Append("], ")

                json.Append("""birthdays"": [ ")
                For Each birthday As DateTime In con.Birthdays
                    json.AppendFormat("""{0}"",", birthday.ToShortDateString)
                Next
                json.Remove(json.Length - 1, 1)
                json.Append("], ")

                json.Append("""websites"": [ ")
                For Each website As String In con.Websites
                    json.AppendFormat("""{0}"",", website)
                Next
                json.Remove(json.Length - 1, 1)
                json.Append("], ")

                json.Append("""companies"": [ ")
                For Each company As ContactCompanyInformation In con.Companies
                    json.AppendFormat("{{""companyname"": ""{0}"", ""jobtitle"": ""{1}""}},", company.CompanyName, company.JobTitle)
                Next
                json.Remove(json.Length - 1, 1)
                json.Append("], ")

                json.Append("""notes"": [ ")
                For Each note As String In con.Notes
                    json.AppendFormat("""{0}"",", note)
                Next
                json.Remove(json.Length - 1, 1)
                json.Append("], ")


                json.AppendLine("},")
            Next
            json.Remove(json.Length - 1, 1)
            json.AppendLine("]}")

            ' Send JSON as an email attachment?

            emailCompose.Subject = "WP JSON contact export"
            emailCompose.Body = json.ToString()
            emailCompose.To = e.Email

            ' MessageBox.Show(json.ToString())

            emailCompose.Show()

        End If


    End Sub


    Private Sub Search_Click(ByVal sender As Object, ByVal e As EventArgs)

        PageTitle.Text = "searching..."
        ContactResultsData.DataContext = Nothing
        Dim cons As Contacts = New Contacts()

        AddHandler cons.SearchCompleted, AddressOf Contacts_SearchCompleted

        cons.SearchAsync("", FilterKind.None, "GetAllContacts")

        'Do work for your application here.
    End Sub

    Private Sub Contacts_SearchCompleted(sender As Object, e As ContactsSearchEventArgs)

        Try
            'Bind the results to the list box that displays them in the UI
            ContactResultsData.DataContext = e.Results

        Catch ex As System.Exception

            'That's okay, no results
        End Try

        If ContactResultsData.Items.Count > 0 Then

            PageTitle.Text = "contacts (" & ContactResultsData.Items.Count & ")"
        Else
            PageTitle.Text = "no contacts!?"
        End If
    End Sub

    Private Sub PhoneApplicationPage_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

    End Sub

    Private Sub ContactResultsData_Tap(sender As System.Object, e As System.Windows.Input.GestureEventArgs) Handles ContactResultsData.Tap
        Dim con As Contact

        con = ContactResultsData.SelectedValue

        If con IsNot Nothing Then
            MessageBox.Show(con.DisplayName)
        End If

    End Sub
End Class
