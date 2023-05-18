<%@ page import="java.util.Map" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="org.springframework.web.multipart.MultipartFile" %>
<%@ page import="java.io.File" %>
<%@ page contentType="text/html; charset=UTF-8" %>
<%@ page pageEncoding="UTF-8"%>
<%
		ArrayList<Map<String, String>> readDateList = (ArrayList<Map<String, String>>) request.getAttribute("sheerMap");
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Insert title here</title>
</head>
<body>
	<div>
		<h2>데이터 파일 구하기</h2>
		<div>
			<button id = 'btn-duplication' class="data-type">수목별 중복수 구하기</button>
			<button class="data-type">흉고 단면적 값 구하기</button>
			<button class="data-type">중요치 구하기</button>
			<button class="data-type">흉고직경급 구하기</button>
		</div>
	</div>
	<div>
		<h2>엑셀 시트 명</h2>
		<% for(int i = 0; i < readDateList.size(); i++ ) { %>
			<button data-value='<%=i%>' class="sheet-name"><%=readDateList.get(i).get(String.valueOf(i))%></button>
		<% } %>
	</div>
	<form id="duplicationForm" enctype="multipart/form-data">
		<input type="file" name="file" />
		<input type="text" name="sheetIndex" id="sheetIndex" value="1"/>
	</form>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
<script type="text/javascript">

	document.getElementById('btn-duplication').onclick = () => {
		var form = $('#duplicationForm')[0];
		var frmData = new FormData(form);
		$.ajax({
			enctype: 'multipart/form-data',
			type : 'post',           // 타입 (get, post, put 등등)
			url : '/excel/duplication.dor',           // 요청할 서버url
			processData: false,
			contentType: false,
			cache: false,
			data: frmData,
			success: function(data) {
				console.log(data);
			},
			error: function(e) {
				console.log(e);
				alert('파일업로드 실패');
			}
		});
	}
</script>

</body>


</html>