<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub ImageButton1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        AccessDataSource1.InsertParameters("�m�W").DefaultValue = NameBox.Text
        AccessDataSource1.InsertParameters("Mail").DefaultValue = MailBox.Text
        AccessDataSource1.InsertParameters("��ĳ").DefaultValue = SgBox.Text
        AccessDataSource1.InsertParameters("���").DefaultValue = ShowDownList.Text
        AccessDataSource1.InsertParameters("�p��").DefaultValue = CoDownList.Text
        AccessDataSource1.InsertParameters("�D�D").DefaultValue = SgDownList.Text

        
        '�ϥ� AccessDataSource1 �g�J�q����Ӹ�Ʈw
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
    <title>���ܤj�n�� ASK CLT</title>
    <link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body>
    
    <div id="wrapper" class="wrapper">
	<div id="top_section" class="top_section">
		<div id="msd_logo" class="msd_logo" title="MSD Be well"></div>
		<div id="top_menu" class="top_menu">
			<ul class="top_menu_item">
				<li class="active"><a href="default.aspx"><img alt="���ܤj�n�� ASK CLT" title="���ܤj�n�� ASK CLT" src="images/msd_topmenu_01.png" /></a></li>
				<li><a href="general.aspx"><img alt="�@����D General Question" title="�@����D General Question" src="images/msd_topmenu_02.png" /></a></li>
				<li><a href="historical.aspx"><img alt="���v���D Historical Data" title="���v���D Historical Data" src="images/msd_topmenu_03.png" /></a></li>
				<li><a href="fornew.aspx"><img alt="����X�� For New MSD" title="����X�� For New MSD" src="images/msd_topmenu_04.png" /></a></li>
			</ul>
		</div>
	</div>
	<div id="hero" class="hero" title="���ܤj�n��  ASK CLT - ���� MSD ���j�p�ơA���A�ӵo�n�ILet us hear you �V Your Voice, We Care!"></div>
	<div id="mid_menu" class="mid_menu">
		<a href="general.aspx"><img alt="�@����D General Question" title="�@����D General Question" src="images/msd_menu_btn_01.png" /></a><a href="historical.aspx"><img alt="���v���D Historical Data" title="���v���D Historical Data" src="images/msd_menu_btn_02.png" /></a><a href="fornew.aspx"><img alt="����X�� For New MSD" title="����X�� For New MSD" src="images/msd_menu_btn_03.png" /></a>
	</div>
	<div id="content_section" class="content_section">
		<div id="welcome_cont_middle" class="welcome_cont_middle">
		<div id="welcome_cont_top" class="welcome_cont_top">
		<div id="welcome_cont_bottom" class="welcome_cont_bottom">
			<div id="welcome_cont" class="welcome_cont">
				<p><img alt="�w��Ө즳�ܤj�n�� Welcome to Ask CLT" title="�w��Ө즳�ܤj�n�� Welcome to Ask CLT" src="images/msd_title_welcome.png" /></p>
				<!--
				<p>�����`���u�@�A�A�����Ī��N���Q�P�D�ޭ̤��ɶܡH<br />���󤽥q���F���A���Ǥ��F�Ѫ��a��ݭn�D�޸Ѵb�ܡH</p>
				<p>�A���N���P���D�A�� MSD Taiwan �ܭ��n�I<br />�uAsk CLT �V ���ܤj�n���v���ѧA�@�ӺZ�ұ������N����y���x�A���׬O�Q�����B�Q�ݪ��A���i�H�b�o�̥R�����F�A���޲z�ζ���F�ѧA���ݭn�P�Q�k�C���קA����ĳ�ΰ��D�O�_��Y�ɳQ�ѨM�A�ڭ̳��D�`�P�§A�����ɡA�åB�@�w�|�J�ӷr�u�A�����n�A�ǤJ�ѦҡC�o�ӷN���H�c���x�� CLT �M�d�޲z�A�ڭ̷|�N�A�����D / �N����浹���������B�z�C</p>
				<p>���ڭ̤@�_�P�ߨ�O���� MSD ��n�I</p>
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
				<p><img alt="�бN�z�_�Q���N����J�U�C��� (Please fill the form below to forward your questions/suggestions)" title="�бN�z�_�Q���N����J�U�C��� (Please fill the form below to forward your questions/suggestions)" src="images/msd_title_form.png" /></p>
			</div>            
            <table cellpadding="0" cellspacing="0" border="0" style="width:100%;">
				<tr>
                    <td style="text-align:right; font-weight:bold; vertical-align:top;">�m�W (Name)�G</td>
					<td style="vertical-align:top; padding-bottom:8px;"><asp:TextBox ID="NameBox" 
                            runat="server" Height="24px" Width="610px"></asp:TextBox></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;">�q�l�l�� (Email)�G</td>
					<td style="vertical-align:top; padding-bottom:8px;"><asp:TextBox ID="MailBox" 
                            runat="server" Height="24px" Width="610px"></asp:TextBox></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;">��ĳ/���D (Suggestion/Comments)�G</td>
					<td style="vertical-align:top; padding-bottom:8px;"><asp:TextBox ID="SgBox" 
                            runat="server" Height="64px" Width="610px"></asp:TextBox><span style="color:red;">*</span></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;"><span style="padding-right:16px;">��ĳ/���D�D�D</span><br />(This suggestions is targetted to)�G</td>
					<td style="vertical-align:top; padding-bottom:8px; padding-top:10px;">
                        <asp:DropDownList ID="SgDownList" runat="server" Height="24px" Width="610px">
                        <asp:ListItem>�@����D</asp:ListItem>
                        <asp:ListItem>����X��</asp:ListItem>
                        </asp:DropDownList></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;"><span style="padding-right:16px;">�s���覡</span><br />(Please Contact me via)�G</td>
					<td style="vertical-align:top; padding-bottom:8px; padding-top:10px;">
                    <asp:DropDownList ID="CoDownList" runat="server" Height="24px" Width="610px">
                        <asp:ListItem>No Contact</asp:ListItem>
                        <asp:ListItem>By Email</asp:ListItem>
                        <asp:ListItem>By Phone</asp:ListItem>
                        <asp:ListItem>Face to Face</asp:ListItem>
                    </asp:DropDownList><span style="color:red;">*</span></td>
				</tr>
				<tr>
					<td style="text-align:right; font-weight:bold; vertical-align:top;"><span style="padding-right:16px;">��ܤ覡</span><br />(I would like this suggestion to be)�G</td>
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
        InsertCommand="INSERT INTO SuggestionData([Suggestion / Comments:], Name, Email, ContactMe, [Public], Question_Type) VALUES (@��ĳ, @�m�W, @Mail, @�p��, @���, @�D�D)"         
        
        
        SelectCommand="SELECT �ѧO�X, Name, Email, [Suggestion / Comments:], [Targetted to], ContactMe, [Public], CreatedDate, ModifyDate, Status, �B�z�Ƶ�, Answer, Post, Question_Type, Question_SubType FROM SuggestionData">
            <InsertParameters>
                <asp:Parameter Name="�m�W" />
                <asp:Parameter Name="Mail" />
                <asp:Parameter Name="��ĳ" />
                <asp:Parameter Name="���" />
                <asp:Parameter Name="�p��" />
                <asp:Parameter Name="�D�D" />

            </InsertParameters>
        </asp:AccessDataSource>
        </asp:View>
        <asp:View ID="View2" runat="server">
			<div id="Div1" class="welcome_cont">
				<p>�z���N���P���D�w�g���\����I<br /></p>
				<p>�P�§A�����ɡI<br /></p>
				
			</div>
        </asp:View>
    </asp:MultiView>&nbsp;
    
    
	<div id="footer_section" class="footer_section">
		<a href="http://www.msd.com.tw/content/corporate/global/privacy_policy.html" target="_blank"><img alt="���p�F���n��" title="���p�F���n��" src="images/msd_footer_01.png" /></a><a href="http://www.msd.com.tw/content/corporate/global/terms_of_use.html" target="_blank"><img alt="�ϥα���" title="�ϥα���" src="images/msd_footer_02.png" /></a><a href="http://www.msd.com.tw/content/corporate/global/copyright.html" target="_blank"><img alt="COPYRIGHT 1995-2011 MERCK & CO., INC." title="COPYRIGHT 1995-2011 MERCK & CO., INC." src="images/msd_footer_03.png" /></a><a href="http://www.merck.com/index.html" target="_blank"><img alt="MERCK & CO., INC. (USA)" title="MERCK & CO., INC. (USA)" src="images/msd_footer_04.png" /></a>
	</div>    
    
    </div>
    </form>


    



</body>
</html>