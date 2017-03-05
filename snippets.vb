Public Sub ShowTombstones()
	Dim de As New DirectoryEntry(My.Settings.oLDAP)
	Dim Searcher As New DirectorySearcher(de)
	'Set our search properties to find Tombstoned objects
	Searcher.PageSize = 1000
	Searcher.Tombstone = True

	'Set the search filter to only find deleted computer objects
	Searcher.Filter = ("(&(isDeleted=TRUE)(objectClass=computer))")
	'Loop through the results and show each deleted object's name
	For Each DeletedObject As SearchResult In Searcher.FindAll
		Dim listObjects As New ListViewItem(New String() {DeletedObject.Properties("name")(0).ToString})
		Form1.lstvwAdminTab.Items.Add(listObjects)

	Next
	Form1.lstvwAdminTab.Sorting = SortOrder.Ascending
End Sub