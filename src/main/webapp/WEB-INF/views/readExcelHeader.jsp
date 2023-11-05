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
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Insert title here</title>
</head>
<body>
	<div>
		<h2>데이터 파일 구하기</h2>
		<div>
			<button id = 'btn-duplication' class="data-type" data-type = '1'>수목별 중복수 구하기</button>
			<button id = 'btn-duplication2' class="data-type" data-type = '2'>흉고 단면적 값 구하기</button>
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
	<form id="duplicationForm" enctype="multipart/form-data" action="/excel/duplication.dor">
		<input type="file" name="file" id="file"/>
		<input type="text" name="sheetIndex" id="sheetIndex" value="1"/>
		<input type="text" name="type" id="type"/>
		<button>azz</button>
	</form>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
<script type="text/javascript">

	document.getElementById('btn-duplication').onclick = (e) => {
		if ('' == document.getElementById('file').value) {
			alert('파일을 선택해주세요')
		} else {
			getExcelFile(e.target.dataset.type);
		}

	}
	document.getElementById('btn-duplication2').onclick = (e) => {
		if ('' == document.getElementById('file').value) {
			alert('파일을 선택해주세요')
		} else {
			getExcelFile(e.target.dataset.type);
		}

	}

	const getExcelFile = (type) => {
		var form = $('#duplicationForm')[0];
		document.getElementById('type').value = type;
		var frmData = new FormData(form);

		$.ajax({
			enctype: 'multipart/form-data',
			type : 'post',           // 타입 (get, post, put 등등)
			url : '/excel/duplication.dor',           // 요청할 서버url
			processData: false,
			contentType: false,
			cache: false,
			data: frmData,
			xhrFields: {
				responseType: "blob",
			}
		})
				.done(function (blob, status, xhr) {
					// check for a filename
					var fileName = "";
					var disposition = xhr.getResponseHeader("Content-Disposition");

					if (disposition && disposition.indexOf("attachment") !== -1) {
						var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
						var matches = filenameRegex.exec(disposition);

						if (matches != null && matches[1]) {
							fileName = decodeURI(matches[1].replace(/['"]/g, ""));
						}
					}

					// for IE
					if (window.navigator && window.navigator.msSaveOrOpenBlob) {
						window.navigator.msSaveOrOpenBlob(blob, fileName);
					} else {
						var URL = window.URL || window.webkitURL;
						var downloadUrl = URL.createObjectURL(blob);

						if (fileName) {
							var a = document.createElement("a");

							// for safari
							if (a.download === undefined) {
								window.location.href = downloadUrl;
							} else {
								a.href = downloadUrl;
								a.download = fileName;
								document.body.appendChild(a);
								a.click();
							}
						} else {
							window.location.href = downloadUrl;
						}
					}
				});
	}
</script>

</body>


</html>