<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeFile="ExcelDownload.aspx.cs" Inherits="ExcelDownload" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="Server">
    <br />
    <div class="page-header">
        <h1>Excel File Download</h1>
    </div>
    <div class="table-responsive">
        <table class="table">
            <tr>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Top 10 News" ID="btnTop10News" OnClick="btnTop10News_Click" />
                </td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Story ROI" ID="btnStoryROI" OnClick="btnStoryROI_Click" />
                </td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Genre ROI" ID="btnGenreROI" OnClick="btnGenreROI_Click" />
                </td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Prog ROI" ID="btnProgROI" OnClick="btnProgROI_Click" />
                </td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Logistics ROI" ID="btnLogisticsROI" OnClick="btnLogisticsROI_Click" />
                </td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Telecast Format ROI" ID="btnTelecastFormatROI" OnClick="btnTelecastFormatROI_Click" />
                </td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Anchor ROI" ID="btnAnchorROI" OnClick="btnAnchorROI_Click" />
                </td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Exclusive Story" ID="btnExclusive" OnClick="btnExclusive_Click" />
                </td>
                <td>
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="UR Split RoI" ID="btnURSplitRoI" OnClick="btnURSplitRoI_Click" />
                </td>
                <td style="display:none;">
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="DoD Story Trend" ID="btnDoDStoryTrend" OnClick="btnDoDStoryTrend_Click" />
                </td>
                <td style="display: none;">
                    <asp:Button runat="server" CssClass="btn btn-primary" Text="Daily Top 10 Story" ID="Button2" OnClick="btnURSplitRoI_Click" />
                </td>
            </tr>
        </table>
    </div>
    <br />
    <div class="table-responsive">
        <asp:GridView runat="server" ID="gvUploadFiles" CssClass="table" AutoGenerateColumns="true">
        </asp:GridView>
    </div>
</asp:Content>

