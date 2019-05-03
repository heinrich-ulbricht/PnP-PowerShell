using System;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using SharePointPnP.PowerShell.CmdletHelpAttributes;
using SharePointPnP.PowerShell.Commands.Base.PipeBinds;

namespace SharePointPnP.PowerShell.Commands.Lists
{
    [Cmdlet(VerbsCommon.Get, "PnPListItem", DefaultParameterSetName = ParameterSet_ALLITEMS)]
    [CmdletHelp("Retrieves list items",
        Category = CmdletHelpCategory.Lists,
        OutputType = typeof(ListItem),
        OutputTypeLink = "https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listitem.aspx")]
    [CmdletExample(
        Code = "PS:> Get-PnPListItem -List Tasks",
        Remarks = "Retrieves all list items from the Tasks list",
        SortOrder = 1)]
    [CmdletExample(
        Code = "PS:> Get-PnPListItem -List Tasks -Id 1",
        Remarks = "Retrieves the list item with ID 1 from the Tasks list",
        SortOrder = 2)]
    [CmdletExample(
        Code = "PS:> Get-PnPListItem -List Tasks -UniqueId bd6c5b3b-d960-4ee7-a02c-85dc6cd78cc3",
        Remarks = "Retrieves the list item with unique id bd6c5b3b-d960-4ee7-a02c-85dc6cd78cc3 from the tasks lists",
        SortOrder = 3)]
    [CmdletExample(
        Code = "PS:> (Get-PnPListItem -List Tasks -Fields \"Title\",\"GUID\").FieldValues",
        Remarks = "Retrieves all list items, but only includes the values of the Title and GUID fields in the list item object",
        SortOrder = 4)]
    [CmdletExample(
        Code = "PS:> Get-PnPListItem -List Tasks -Query \"<View><Query><Where><Eq><FieldRef Name='GUID'/><Value Type='Guid'>bd6c5b3b-d960-4ee7-a02c-85dc6cd78cc3</Value></Eq></Where></Query></View>\"",
        Remarks = "Retrieves all list items based on the CAML query specified",
        SortOrder = 5)]
    [CmdletExample(
        Code = "PS:> Get-PnPListItem -List Tasks -PageSize 1000",
        Remarks = "Retrieves all list items from the Tasks list in pages of 1000 items",
        SortOrder = 6)]
    [CmdletExample(
        Code = "PS:> Get-PnPListItem -List Tasks -PageSize 1000 -ScriptBlock { Param($items) $items.Context.ExecuteQuery() } | % { $_.BreakRoleInheritance($true, $true) }",
        Remarks = "Retrieves all list items from the Tasks list in pages of 1000 items and breaks permission inheritance on each item",
        SortOrder = 7)]
    public class GetListItem : PnPWebCmdlet
    {
        private const string ParameterSet_BYID = "By Id";
        private const string ParameterSet_BYUNIQUEID = "By Unique Id";
        private const string ParameterSet_BYQUERY = "By Query";
        private const string ParameterSet_ALLITEMS = "All Items";
        [Parameter(Mandatory = true, ValueFromPipeline = true, HelpMessage = "The list to query", Position = 0, ParameterSetName = ParameterAttribute.AllParameterSets)]
        public ListPipeBind List;

        [Parameter(Mandatory = false, HelpMessage = "The ID of the item to retrieve", ParameterSetName = ParameterSet_BYID)]
        public int Id = -1;

        [Parameter(Mandatory = false, HelpMessage = "The unique id (GUID) of the item to retrieve", ParameterSetName = ParameterSet_BYUNIQUEID)]
        public GuidPipeBind UniqueId;

        [Parameter(Mandatory = false, HelpMessage = "The CAML query to execute against the list", ParameterSetName = ParameterSet_BYQUERY)]
        public string Query;

        [Parameter(Mandatory = false, HelpMessage = "The fields to retrieve. If not specified all fields will be loaded in the returned list object.", ParameterSetName = ParameterSet_ALLITEMS)]
        [Parameter(Mandatory = false, HelpMessage = "The fields to retrieve. If not specified all fields will be loaded in the returned list object.", ParameterSetName = ParameterSet_BYID)]
        [Parameter(Mandatory = false, HelpMessage = "The fields to retrieve. If not specified all fields will be loaded in the returned list object.", ParameterSetName = ParameterSet_BYUNIQUEID)]
        public string[] Fields;

        [Parameter(Mandatory = false, HelpMessage = "The number of items to retrieve per page request.", ParameterSetName = ParameterSet_ALLITEMS)]
		[Parameter(Mandatory = false, HelpMessage = "The number of items to retrieve per page request.", ParameterSetName = ParameterSet_BYQUERY)]
        public int PageSize = -1;

		[Parameter(Mandatory = false, HelpMessage = "The script block to run after every page request.", ParameterSetName = ParameterSet_ALLITEMS)]
		[Parameter(Mandatory = false, HelpMessage = "The script block to run after every page request.", ParameterSetName = ParameterSet_BYQUERY)]
		public ScriptBlock ScriptBlock;

		protected override void ExecuteCmdlet()
        {
            var list = List.GetList(SelectedWeb);

            if (HasId())
            {
                var listItem = list.GetItemById(Id);
                if (Fields != null)
                {
                    foreach (var field in Fields)
                    {
                        ClientContext.Load(listItem, l => l[field]);
                    }
                }
                else
                {
                    ClientContext.Load(listItem);
                }
                ClientContext.ExecuteQueryRetry();
                WriteObject(listItem);
            }
            else if (HasUniqueId())
            {
                CamlQuery query = new CamlQuery();
                var viewFieldsStringBuilder = new StringBuilder();
                if (HasFields())
                {
                    viewFieldsStringBuilder.Append("<ViewFields>");
                    foreach (var field in Fields)
                    {
                        viewFieldsStringBuilder.AppendFormat("<FieldRef Name='{0}'/>", field);
                    }
                    viewFieldsStringBuilder.Append("</ViewFields>");
                }
                query.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='GUID'/><Value Type='Guid'>{UniqueId.Id}</Value></Eq></Where></Query>{viewFieldsStringBuilder}</View>";
                var listItem = list.GetItems(query);
                ClientContext.Load(listItem);
                ClientContext.ExecuteQueryRetry();
                WriteObject(listItem);
            }
            else
            {
				CamlQuery query = HasCamlQuery() ? new CamlQuery { ViewXml = Query } : CamlQuery.CreateAllItemsQuery();

				if (Fields != null)
                {
                    var queryElement = XElement.Parse(query.ViewXml);

                    var viewFields = queryElement.Descendants("ViewFields").FirstOrDefault();
                    if (viewFields != null)
                    {
                        viewFields.RemoveAll();
                    }
                    else
                    {
                        viewFields = new XElement("ViewFields");
                        queryElement.Add(viewFields);
                    }

                    foreach (var field in Fields)
                    {
                        XElement viewField = new XElement("FieldRef");
                        viewField.SetAttributeValue("Name", field);
                        viewFields.Add(viewField);
                    }
                    query.ViewXml = queryElement.ToString();
                }

                if (HasPageSize())
                {
                    var queryElement = XElement.Parse(query.ViewXml);

                    var rowLimit = queryElement.Descendants("RowLimit").FirstOrDefault();
                    if (rowLimit != null)
                    {
                        rowLimit.RemoveAll();
                    }
                    else
                    {
                        rowLimit = new XElement("RowLimit");
                        queryElement.Add(rowLimit);
                    }

                    rowLimit.SetAttributeValue("Paged", "TRUE");
                    rowLimit.SetValue(PageSize);

                    query.ViewXml = queryElement.ToString();
                }

                try
                {
                    do
                    {
                        var listItems = list.GetItems(query);
                        ClientContext.Load(listItems);
                        ClientContext.ExecuteQueryRetry();

                        WriteObject(listItems, true);

                        if (ScriptBlock != null)
                        {
                            ScriptBlock.Invoke(listItems);
                        }

                        query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                    } while (query.ListItemCollectionPosition != null);
                } catch (ServerException e)
                {
                    if (e.ServerErrorCode == -2147024860 && e.ServerErrorTypeName == "Microsoft.SharePoint.SPQueryThrottledException")
                    {
                        // check if we can use "special" iteration logic that works by paging over the ID (this works since ID is indexed)
                        if (!HasCamlQuery())
                        {
                            // first get the maximum ID for the list to know where paging ends
                            var originalViewXml = query.ViewXml;
                            query.ViewXml = @"<View> <Query> <OrderBy> <FieldRef Name='ID' Ascending='FALSE' /> </OrderBy> </Query> <ViewFields> <FieldRef Name='Id' /> </ViewFields> <RowLimit>1</RowLimit></View>";
                            var listItems = list.GetItems(query);
                            ClientContext.Load(listItems);
                            ClientContext.ExecuteQueryRetry();
                            if (listItems.Count != 1)
                            {
                                // something went wrong - just throw the original exception
                                throw;
                            }
                            var maxId = listItems[0].Id;
                            var currentPage = 0;
                            // use specified page size or a default near the limit
                            var pageSize = PageSize;
                            if (pageSize <= 0)
                            {
                                pageSize = 4999;
                            }
                            do
                            {
                                var queryElement = XElement.Parse(originalViewXml);
                                var rowLimit = queryElement.Descendants("RowLimit").FirstOrDefault();
                                if (rowLimit != null)
                                {
                                    rowLimit.RemoveAll();
                                }
                                var queryChildElement = queryElement.Descendants("Query").FirstOrDefault();
                                if (queryChildElement != null)
                                {
                                    // build paged query
                                    var whereElement = XElement.Parse($@"<Where><And><Gt><FieldRef Name='ID' /><Value Type='Text'>{currentPage * pageSize}</Value></Gt><Lt><FieldRef Name='ID' /><Value Type='Text'>{(currentPage + 1) * pageSize}</Value></Lt></And></Where>");
                                    queryChildElement.Add(whereElement);
                                }

                                query.ViewXml = queryElement.ToString();
                                listItems = list.GetItems(query);
                                ClientContext.Load(listItems);
                                ClientContext.ExecuteQueryRetry();

                                WriteObject(listItems, true);

                                if (ScriptBlock != null)
                                {
                                    ScriptBlock.Invoke(listItems);
                                }

                                currentPage++;
                            } while (currentPage * pageSize < maxId);
                        } else
                        {
                            // the user specified a CAML query? can't handle this
                            throw;
                        }
                    }
                    else
                    {
                        throw;
                    }
                    
                }
            }
        }

        private bool HasId()
        {
            return Id != -1;
        }

        private bool HasUniqueId()
        {
            return UniqueId != null && UniqueId.Id != Guid.Empty;
        }

        private bool HasCamlQuery()
        {
            return Query != null;
        }

        private bool HasFields()
        {
            return Fields != null;
        }

        private bool HasPageSize()
        {
            return PageSize > 0;
        }
    }
}
