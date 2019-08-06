<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta content="text/html; charset=big5" http-equiv="Content-Type" />
<title>有話大聲說 ASK CLT | 一般問題 General Question</title>
<link href="style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="jquery-1.7.1.min.js"></script>

<script type="text/javascript">
    $(document).ready(function () {
        $('#info tr').addClass('odd');
        $('#info tr:even').addClass('even');
    });
</script>
</head>
<body>

<div id="wrapper" class="wrapper">
	<div id="top_section" class="top_section">
		<div id="msd_logo" class="msd_logo" title="MSD Be well"></div>
		<div id="top_menu" class="top_menu">
			<ul class="top_menu_item">
				<li><a href="default.aspx"><img alt="有話大聲說 ASK CLT" title="有話大聲說 ASK CLT" src="images/msd_topmenu_01.png" /></a></li>
				<li class="active"><a href="general.aspx"><img alt="一般問題 General Question" title="一般問題 General Question" src="images/msd_topmenu_02.png" /></a></li>
				<li><a href="historical.aspx"><img alt="歷史問題 Historical Data" title="歷史問題 Historical Data" src="images/msd_topmenu_03.png" /></a></li>
				<li><a href="fornew.aspx"><img alt="關於合併 For New MSD" title="關於合併 For New MSD" src="images/msd_topmenu_04.png" /></a></li>
			</ul>
		</div>
	</div>
	<div id="hero" class="hero" title="有話大聲說  ASK CLT - 關於 MSD 的大小事，等你來發聲！Let us hear you – Your Voice, We Care!"></div>
	<div id="qanda_menu_section" class="qanda_menu_section">
		<div id="qanda_menu" class="qanda_menu">
			<img alt="一般問題 (General Question)" title="一般問題 (General Question)" src="images/msd_qandamenu_01on.png" /><a href="historical.aspx"><img alt="歷史問題 (Historical Data)" title="歷史問題 (Historical Data)" src="images/msd_qandamenu_02.png" /></a><a href="fornew.aspx"><img alt="關於合併 (For New MSD)" title="關於合併 (For New MSD)" src="images/msd_qandamenu_03.png" /></a>
		</div>
        <!-- <div id="qanda_search" class="qanda_search">
			<div id="searchBox" class="searchBox">
				<form>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td><input type="text" name="strSearchTerm" class="text" maxlength="100" value=""/></td>
							<td><input type="image" src="images/butn_search.gif" class="butn" /></td>
						</tr>
					</table>
				</form>
			</div>
		</div> -->
	</div>

    <form id="form2" runat="server">
    <asp:AccessDataSource ID="AccessDataSource1" runat="server" 
        DataFile="~/App_Data/CLT_MailBox.mdb" 
        
        
        
        
        
        SelectCommand="SELECT [Suggestion / Comments:] AS Suggestion, [Answer], format(CreatedDate, 'yyyy/mm/dd') AS CreatedDate FROM [SuggestionData] WHERE Question_Type = '一般問題' AND Post = -1 ORDER BY [CreatedDate] DESC">
        <SelectParameters>
            <asp:Parameter DefaultValue="260" Name="識別碼" Type="Int32" />
            <asp:Parameter DefaultValue="280" Name="識別碼2" Type="Int32" />
        </SelectParameters>
    </asp:AccessDataSource>

    <div id="Div1" class="content_section">
    
        <asp:ListView ID="ListView1" runat="server" 
            DataSourceID="AccessDataSource1">
            
            <AlternatingItemTemplate>
                <tr style="">
                    <td class="date">
                        <asp:Label ID="CreatedDateLabel" runat="server" 
                            Text='<%# Eval("CreatedDate") %>' />
                    </td>                    
                    <td class="question">
                        <asp:Label ID="SuggestionLabel" runat="server" Text='<%# Eval("Suggestion") %>' />
                    </td>
                    <td class="answer">
                        <asp:Label ID="AnswerLabel" runat="server" Text='<%# Eval("Answer") %>' />
                    </td>

                </tr>
            </AlternatingItemTemplate>
            
            <EditItemTemplate>
                <tr style="">
                    <td>
                        <asp:Button ID="UpdateButton" runat="server" CommandName="Update" 
                            Text="Update" />
                        <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" 
                            Text="Cancel" />
                    </td>
                    <td>
                        <asp:TextBox ID="CreatedDateTextBox" runat="server" 
                            Text='<%# Bind("CreatedDate") %>' />
                    </td>                    
                    <td>
                        <asp:TextBox ID="SuggestionTextBox" runat="server" Text='<%# Bind("Suggestion") %>' />
                    </td>
                    <td>
                        <asp:TextBox ID="AnswerTextBox" runat="server" Text='<%# Bind("Answer") %>' />
                    </td>

                </tr>
            </EditItemTemplate>
            
            <EmptyDataTemplate>
                <table id="Table1" runat="server" style="">
                    <tr>
                        <td>
                            No data was returned.</td>
                    </tr>
                </table>
            </EmptyDataTemplate>
            
            <InsertItemTemplate>
                <tr style="">
                    <td>
                        <asp:Button ID="InsertButton" runat="server" CommandName="Insert" 
                            Text="Insert" />
                        <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" 
                            Text="Clear" />
                    </td>
                    <td>
                        <asp:TextBox ID="CreatedDateTextBox" runat="server" 
                            Text='<%# Bind("CreatedDate") %>' />
                    </td>                    
                    <td>
                        <asp:TextBox ID="SuggestionTextBox" runat="server" Text='<%# Bind("Suggestion") %>' />
                    </td>
                    <td>
                        <asp:TextBox ID="AnswerTextBox" runat="server" Text='<%# Bind("Answer") %>' />
                    </td>

                </tr>
            </InsertItemTemplate>
            
            <ItemTemplate>
                <tr style="">
                    <td class="date">
                        <asp:Label ID="CreatedDateLabel" runat="server" 
                            Text='<%# Eval("CreatedDate") %>' />
                    </td>                    
                    <td class="question">
                        <asp:Label ID="SuggestionLabel" runat="server" Text='<%# Eval("Suggestion") %>' />
                    </td>
                    <td class="answer">
                        <asp:Label ID="AnswerLabel" runat="server" Text='<%# Eval("Answer") %>' />
                    </td>

                </tr>
            </ItemTemplate>
            
            <LayoutTemplate>
                <table id="Table1"  runat="server">
                    <tr id="Tr1" runat="server">
                        <td id="Td1" runat="server">
                            <table cellpadding="0" cellspacing="0" class="info"  ID="itemPlaceholderContainer" runat="server" border="0" style="">
                                <tr id="Tr2" runat="server" style="">                                    
                                    <th id="Th1" runat="server">
                                        <img alt="日期 (Date)" title="日期 (Date)" src="images/msd_qanda_table_date.jpg" /></th>                                    
                                    <th id="Th2" runat="server">
                                        <img alt="問題 (Question)" title="問題 (Question)" src="images/msd_qanda_table_question.jpg" /></th>
                                    <th id="Th3" runat="server">
                                        <img alt="答案 (Answer)" title="答案 (Answer)" src="images/msd_qanda_table_answer.jpg" /></th>

                                </tr>
                                <tr ID="itemPlaceholder" runat="server">
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr id="Tr3" runat="server">
                        <td id="Td2" runat="server" style="">
                        </td>
                    </tr>
                </table>
            </LayoutTemplate>
            
            <SelectedItemTemplate>
                <tr style="">
                    <td>
                        <asp:Label ID="CreatedDateLabel" runat="server" 
                            Text='<%# Eval("CreatedDate") %>' />
                    </td>                    
                    <td>
                        <asp:Label ID="SuggestionLabel" runat="server" Text='<%# Eval("Suggestion") %>' />
                    </td>
                    <td>
                        <asp:Label ID="AnswerLabel" runat="server" Text='<%# Eval("Answer") %>' />
                    </td>

                </tr>
            </SelectedItemTemplate>

        </asp:ListView>
    
    </div>
    </form>

	<div id="footer_section" class="footer_section">
		<a href="http://www.msd.com.tw/content/corporate/global/privacy_policy.html" target="_blank"><img alt="隱私政策聲明" title="隱私政策聲明" src="images/msd_footer_01.png" /></a><a href="http://www.msd.com.tw/content/corporate/global/terms_of_use.html" target="_blank"><img alt="使用條款" title="使用條款" src="images/msd_footer_02.png" /></a><a href="http://www.msd.com.tw/content/corporate/global/copyright.html" target="_blank"><img alt="COPYRIGHT 1995-2011 MERCK & CO., INC." title="COPYRIGHT 1995-2011 MERCK & CO., INC." src="images/msd_footer_03.png" /></a><a href="http://www.merck.com/index.html" target="_blank"><img alt="MERCK & CO., INC. (USA)" title="MERCK & CO., INC. (USA)" src="images/msd_footer_04.png" /></a>
	</div>
</div>
</body>
</html>
