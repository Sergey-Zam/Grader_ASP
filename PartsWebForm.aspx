<%@ Page Title="Сравнение деталей" Language="vb" Async="true" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="PartsWebForm.aspx.vb" Inherits="Grader_ASP.PartsWebForm" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <!--Начало контента-->
    <div class="right_div">
        <h3>Сравнение деталей</h3>

        <div class="line_div">
            Файл эталонной детали (.ipt) <br />
            <input type=file id="locationOfStandardDocument" runat="server" accept=".ipt" size="50"/>  <br /> 
            Файл проверяемой детали (.ipt) <br />      
            <input type="file" id="locationOfCheckedDocument" runat="server" accept=".ipt" size="50"/> <br /> 
            <input type="submit" class="subButton" id="SubmitToServer" value="Загрузить файлы на сервер и получить результат" runat="server" onclick="SubmitToServer_Click" style="width: 100%;"/>
        </div>
        <div class="line_div" style="text-align: center;" >       
            Таблица результатов
            <!--таблица результатов-->
            <div id="tableOfResults" class="tableOfResults" runat="server"></div>
            <!--подпись под таблицей-->
            <div>
               Всего получено строк:
               <asp:Label ID="lblCountOfRows" runat="server" Text="0"></asp:Label>
            </div> 
        </div>
        <div class="line_div" style="text-align: right;">            
            <asp:Button ID="btnExportTable" runat="server" Text="Экспорт таблицы..." BackColor="White" BorderStyle="Solid" ForeColor="#1565C0" BorderColor="#1565C0" BorderWidth="2px" />
            <asp:Button ID="btnClearTable" runat="server" style="margin-left: 5px" Text="Очистить таблицу" BackColor="White" BorderStyle="Solid" ForeColor="#1565C0" BorderColor="#1565C0" BorderWidth="2px" />
        </div>
    </div>
    <!--Конец контента-->
</asp:Content>
