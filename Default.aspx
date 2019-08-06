<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub ImageButton1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        AccessDataSource1.InsertParameters("姓名").DefaultValue = NameBox.Text
        AccessDataSource1.InsertParameters("Mail").DefaultValue = MailBox.Text
        AccessDataSource1.InsertParameters("建議").DefaultValue = SgBox.Text
        AccessDataSource1.InsertParameters("顯示").DefaultValue = ShowDownList.Text
        AccessDataSource1.InsertParameters("聯絡").DefaultValue = CoDownList.Text
        AccessDataSource1.InsertParameters("主題").DefaultValue = SgDownList.Text

        
        '使用 AccessDataSource1 寫入訂單明細資料庫
        AccessDataSource1.Insert()
        MultiView1.ActiveViewIndex = 1
    End Sub

    Protected Sub ImageButton2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        NameBox.Text = ""
        MailBox.Text = ""
        SgBox.Text = ""
        
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta content="text/html; charset=big5" http-equiv="Content-Type" />
    <title>有話大聲說 ASK CLT</title>
    <link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body>
    
    <div id="wrapper" class="wrapper">
	<div id="top_section" class="top_section">
		<div id="msd_logo" class="msd_logo" title="MSD Be well"></div>
		<div id="top_menu" class="top_menu">
			<ul class="top_menu_item">
				<li class="active"><a href="default.aspx"><img alt="有話大聲說 ASK CLT" title="有話大聲說 ASK CLT" src="images/msd_topmenu_01.png" /></a></li>
				<li><a href="general.aspx"><img alt="一般問題 General Question" title="一般問題 General Question" src="images/msd_topmenu_02.png" /></a></li>
				<li><a href="historical.aspx"><img alt="歷史問題 Historical Data" title="歷史問題 Historical Data" src="images/msd_topmenu_03.png" /></a></li>
				<li><a href="fornew.aspx"><img alt="關於合併 For New MSD" title="關於合併 For New MSD" src="images/msd_topmenu_04.png" /></a></li>
			</ul>
		</div>
	</div>
	<div id="hero" class="hero" title="有話大聲說  ASK CLT - 關於 MSD 的大小事，等你來發聲！Let us hear you – Your Voice, We Care!"></div>
	<div id="mid_menu" class="mid_menu">
		<a href="general.aspx"><img alt="一般問題 General Question" title="一般問題 General Question" src="images/msd_menu_btn_01.png" /></a><a href="historical.aspx"><img alt="歷史問題 Historical Data" title="歷史問題 Historical Data" src="images/msd_menu_btn_02.png" /></a><a href="fornew.aspx"><img alt="關於合併 For New MSD" title="關於合併 For New MSD" src="images/msd_menu_btn_03.png" /></a>
	</div>
	<div id="content_section" class="content_section">
		<div id="welcome_cont_middle" class="welcome_cont_middle">
		<div id="welcome_cont_top" class="welcome_cont_top">
		<div id="welcome_cont_bottom" class="welcome_cont_bottom">
			<div id="welcome_cont" class="welcome_cont">
				<p><img alt="歡迎來到有話大聲說 Welcome to Ask CLT" title="歡迎來到有話大聲說 Welcome to Ask CLT" src="images/msd_title_welcome.png" /></p>
				<!--
				<p>面對日常的工作，你有滿腔的意見想與主管們分享嗎？<br />關於公司的政策，有些不了解的地方需要主管解惑嗎？</p>
				<p>你的意見與問題，對 MSD Taiwan 很重要！<br />「Ask CLT – 有話大聲說」提供你一個暢所欲言的意見交流平台，不論是想說的、想問的，都可以在這裡充分表達，讓管理團隊更了解你的需要與想法。不論你的建議或問題是否能即時被解決，我們都非常感謝你的分享，並且一定會仔細斟酌你的心聲，納入參考。這個意見信箱平台由 CLT 專責管理，我們會將你的問題 / 意見轉交給相關部門處理。</p>
				<p>讓我們一起同心協力的讓 MSD 更好！</p>
				-->
				<p><img alt="" title="" src="images/msd_con_welcome.png" /></p>
				<p><img alt="" title="" src="images/msd_con_welcome_en.png" /></p>
			</div>
		</div>
		</div>
		</div>
		<div id="askclt_form_section" class="askclt_form_section">

        <form id="form1" runat="server">
        <asp:MultiView ID="MultiView1" runat="server" ActiveViewIndex="0" >    
            <asp:View ID="View1" runat="server">			
			<div id="sub_title" class="sub_title">
				<p><img alt="請將您寶貴的意見填入下列欄位 (Please fill the form below to forward your questions/suggestions)" title="請將您寶貴的意見填入下列欄位 (Please fill the form below to forward your questions/suggestions)" src="images/msd_title_form.png" /></p>
			</div>            
            <table cellpadding="0" cellspacing="0" border="0" style="width:100%;">
				<tr>
                    <td style="text-align:right; font-weight:bold; vertical-align:top;">姓名 (Name)：</td>
					<td style="vertical-align:top; padding-bottom:8px;"><asp:TextBox ID="NameBox" 
                            runat="server" Height="24px" Width="610px"></asp:TextBox></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;">電子郵件 (Email)：</td>
					<td style="vertical-align:top; padding-bottom:8px;"><asp:TextBox ID="MailBox" 
                            runat="server" Height="24px" Width="610px"></asp:TextBox></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;">建議/問題 (Suggestion/Comments)：</td>
					<td style="vertical-align:top; padding-bottom:8px;"><asp:TextBox ID="SgBox" 
                            runat="server" Height="64px" Width="610px"></asp:TextBox><span style="color:red;">*</span></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;"><span style="padding-right:16px;">建議/問題主題</span><br />(This suggestions is targetted to)：</td>
					<td style="vertical-align:top; padding-bottom:8px; padding-top:10px;">
                        <asp:DropDownList ID="SgDownList" runat="server" Height="24px" Width="610px">
                        <asp:ListItem>一般問題</asp:ListItem>
                        <asp:ListItem>關於合併</asp:ListItem>
                        </asp:DropDownList></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;"><span style="padding-right:16px;">連絡方式</span><br />(Please Contact me via)：</td>
					<td style="vertical-align:top; padding-bottom:8px; padding-top:10px;">
                    <asp:DropDownList ID="CoDownList" runat="server" Height="24px" Width="610px">
                        <asp:ListItem>No Contact</asp:ListItem>
                        <asp:ListItem>By Email</asp:ListItem>
                        <asp:ListItem>By Phone</asp:ListItem>
                        <asp:ListItem>Face to Face</asp:ListItem>
                    </asp:DropDownList><span style="color:red;">*</span></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;"><span style="padding-right:16px;">顯示方式</span><br />(I would like this suggestion to be)：</td>
					<td style="vertical-align:top; padding-bottom:8px; padding-top:10px;">
                    <asp:DropDownList ID="ShowDownList" runat="server" Height="24px" Width="610px">
                        <asp:ListItem>Public</asp:ListItem>
                        <asp:ListItem>Confidential</asp:ListItem>
                    </asp:DropDownList><span style="color:red;">*</span></td>
				</tr>
				<tr>
					<td colspan="2" style="padding-top:30px; padding-bottom:6px; background-image:url('images/msd_form_btn_bg.jpg'); background-repeat:no-repeat; background-position:center bottom; text-align:center;">
						    <asp:ImageButton ID="ImageButton1" runat="server" 
                                ImageUrl="~/images/msd_form_btn_submit.png" onclick="ImageButton1_Click" />
                    &nbsp;
                    <asp:ImageButton ID="ImageButton2" runat="server" 
                                ImageUrl="~/images/msd_form_btn_clear.png" onclick="ImageButton2_Click" /><br />
					</td>
				</tr>
			</table>
		</div>
	</div>
        <asp:AccessDataSource ID="AccessDataSource1" runat="server" 
        DataFile="~/App_Data/CLT_MailBox.mdb" 
        InsertCommand="INSERT INTO SuggestionData([Suggestion / Comments:], Name, Email, ContactMe, [Public], Question_Type) VALUES (@建議, @姓名, @Mail, @聯絡, @顯示, @主題)"         
        
        
        SelectCommand="SELECT 識別碼, Name, Email, [Suggestion / Comments:], [Targetted to], ContactMe, [Public], CreatedDate, ModifyDate, Status, 處理備註, Answer, Post, Question_Type, Question_SubType FROM SuggestionData">
            <InsertParameters>
                <asp:Parameter Name="姓名" />
                <asp:Parameter Name="Mail" />
                <asp:Parameter Name="建議" />
                <asp:Parameter Name="顯示" />
                <asp:Parameter Name="聯絡" />
                <asp:Parameter Name="主題" />

            </InsertParameters>
        </asp:AccessDataSource>
        </asp:View>
        <asp:View ID="View2" runat="server">
			<div id="Div1" class="welcome_cont">
				<p>您的意見與問題已經成功提交！<br /></p>
				<p>感謝你的分享！<br /></p>
				
			</div>
        </asp:View>
    </asp:MultiView>&nbsp;
    
    
	<div id="footer_section" class="footer_section">
		<a href="http://www.msd.com.tw/content/corporate/global/privacy_policy.html" target="_blank"><img alt="隱私政策聲明" title="隱私政策聲明" src="images/msd_footer_01.png" /></a><a href="http://www.msd.com.tw/content/corporate/global/terms_of_use.html" target="_blank"><img alt="使用條款" title="使用條款" src="images/msd_footer_02.png" /></a><a href="http://www.msd.com.tw/content/corporate/global/copyright.html" target="_blank"><img alt="COPYRIGHT 1995-2011 MERCK & CO., INC." title="COPYRIGHT 1995-2011 MERCK & CO., INC." src="images/msd_footer_03.png" /></a><a href="http://www.merck.com/index.html" target="_blank"><img alt="MERCK & CO., INC. (USA)" title="MERCK & CO., INC. (USA)" src="images/msd_footer_04.png" /></a>
	</div>    
    
    </div>
    </form>


    



</body>
</html>
