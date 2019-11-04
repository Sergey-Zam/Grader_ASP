<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="MainWebForm.aspx.vb" Inherits="Grader_ASP.MainWebForm" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    </asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">   
    <!--Начало контента-->
    <div class="content_div">
        <div class="left_div">
            <div class="header_div">Grader</div>

            <h2> Алгоритм работы с программой: </h2>
            <h4>1. Импортируйте документ эталонной сборки</h4>
            <h4>2. Импортируйте документ проверяемой сборки</h4>
            <h4>3. Подождите</h4>
            <h4>4. Полученный результат можно экспортировать в формате Excel</h4>           
        </div>
        <div class="right_div">
            <div class="line_div">
                Файл эталонной сборки (.iam) <br />             
                <input type=file id=locationOfStandardAssembly runat="server" accept=".iam" size="50"/>  <br /> 
                Файл проверяемой сборки (.iam) <br />          
                <input type="file" id="locationOfCheckedAssembly" runat="server" accept=".iam" size="50"/> <br /> 
                <input type="submit" id="SubmitToServer" value="Загрузить файлы на сервер и получить результат" runat="server" onclick="SubmitToServer_Click" style="width: 100%;" class="button"/>
            </div>
            <div class="line_div" style="text-align: center;" >               
                Таблица результатов
                <!--таблица результатов-->
                <div id="tableOfResults" class="tableOfResults" runat="server">
                </div>
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
    </div>
    <!--Конец контента-->
</asp:Content>
