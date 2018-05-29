<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeFile="ExcelUpload.aspx.cs" Inherits="ExcelUpload" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="Server">
    <br />
    <div class="page-header">
        <h1>Excel File Upload</h1>
    </div>
    <div class="table-responsive">
        <table class="table">
            <tr>
                <td>File Date</td>
                <td>
                    <asp:TextBox type="date" runat="server" ID="txtDate" CssClass="form-control"></asp:TextBox></td>
            </tr>
            <tr>
                <td>File Name</td>
                <td>
                    <asp:FileUpload ID="fileUpload" type="file" runat="server" name="fileUpload" accept=".xls, .xlsx" CssClass="form-control" /></td>
            </tr>
            <tr>
                <td></td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Save File" ID="btnSubmit" OnClick="btnSubmit_Click" />
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Label ID="lblMessage" Font-Bold="true" runat="server" Font-Size="10" Font-Name="verdana"></asp:Label></td>
            </tr>
        </table>
    </div>
    <br />
    <div class="table-responsive">
        <asp:GridView runat="server" ID="gvUploadFiles" CssClass="table" AutoGenerateColumns="false" OnRowCommand="gvUploadFiles_RowCommand">
            <Columns>
                <asp:TemplateField HeaderText="S.No.">
                    <ItemTemplate>
                        <%# Container.DataItemIndex+1 %>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="File Date">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lbl_Selected_date" Text='<%# Eval("Selected_date").ToString() %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="File Name">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lbl_original_file_name" Text='<%# Eval("original_file_name").ToString() %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="File upload Time">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lbl_upload_date" Text='<%# Eval("upload_date").ToString() %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Sync Status">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lbl_synchronize_status" Text='<%# ((Eval("synchronize_status").ToString()=="0")?"Not Synchronized":((Eval("synchronize_status").ToString()=="1")?"Partial":((Eval("synchronize_status").ToString()=="2")?"Synchronized":""))) %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Sync Date">
                    <ItemTemplate>
                        <asp:LinkButton runat="server" Text="Sync Now" CommandName="SyncNow" CommandArgument='<%# Eval("upload_id") %>' Visible='<%# ((Eval("synchronize_status").ToString()=="2")?false:true) %>'></asp:LinkButton>
                        <asp:Label runat="server" ID="lbl_synchronize_date" Text='<%# Eval("synchronize_date") %>' Visible='<%# ((Eval("synchronize_status").ToString()=="2")?true:false) %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Delete">
                    <ItemTemplate>
                        <asp:LinkButton runat="server" Text="Delete" CommandName="DeleteNow" CommandArgument='<%# Eval("upload_id") %>' Visible='<%# ((Eval("synchronize_status").ToString()=="2")?false:true) %>'></asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </div>
</asp:Content>

