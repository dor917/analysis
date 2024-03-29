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

		<form id="duplicationForm" enctype="multipart/form-data" action='retrun false;'>
			<input type="hidden" name="type" id="type"/>
			<div>
				<input type="file" name="file" id="file"/>
			</div>
			<div>
				<label for="sheetIndex">임목조사표 시트 순번(0부터 시작) :</label>
				<input type="text" name="sheetIndex" id="sheetIndex" value="1">
			</div>
			<div>
				<button id = 'btn-duplication3' type="button" class="data-type" data-type = '3'>중요치 구하기</button>
			</div>
			<div>
				<label for="sheetIndex">흉고직경 구하기 나무명 :</label>
				<input type="text" name="treeName" id="treeName"/>
			</div>
			<div>
				<button id = 'btn-duplication4' type="button" class="data-type" data-type = '4'>흉고직경급 구하기</button>
			</div>
		</form>
	</div>
<%--	<div>--%>
<%--		<h2>엑셀 시트 명</h2>--%>
<%--		<% for(int i = 0; i < readDateList.size(); i++ ) { %>--%>
<%--			<button data-value='<%=i%>' class="sheet-name"><%=readDateList.get(i).get(String.valueOf(i))%></button>--%>
<%--		<% } %>--%>
<%--	</div>--%>




<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
<script type="text/javascript">

	// document.getElementById('btn-duplication').onclick = (e) => {
	// 	if ('' == document.getElementById('file').value) {
	// 		alert('파일을 선택해주세요')
	// 	} else {
	// 		getExcelFile(e.target.dataset.type);
	// 	}
	// }
	// document.getElementById('btn-duplication2').onclick = (e) => {
	// 	if ('' == document.getElementById('file').value) {
	// 		alert('파일을 선택해주세요')
	// 	} else {
	// 		getExcelFile(e.target.dataset.type);
	// 	}
	// }
	document.getElementById('btn-duplication3').onclick = (e) => {
		if ('' == document.getElementById('file').value) {
			alert('파일을 선택해주세요')
		} else {
			getExcelFile(e.target.dataset.type);
		}
	}
	document.getElementById('btn-duplication4').onclick = (e) => {
		if ('' == document.getElementById('file').value) {
			alert('파일을 선택해주세요')
		} else {
			if (document.getElementById('treeName').value != '') {
				getExcelFile('5');
			} else {
				getExcelFile(e.target.dataset.type);
			}

		}
	}
	const getExcelFile = (type) => {
		var form = $('#duplicationForm')[0];
		document.getElementById('type').value = type;
		var frmData = new FormData(form);
		var getUrl = '/excel/duplication.dor';
		if (type == 5) {
			getUrl = '/excel/getDiameterForTreeName.dor';
		}
		var ge = '/excel/duplication.dor';
		$.ajax({
			enctype: 'multipart/form-data',
			type : 'post',           // 타입 (get, post, put 등등)
			url : getUrl,           // 요청할 서버url
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


	const getDiameterForTreeName = (type) => {
		var form = $('#duplicationForm')[0];
		document.getElementById('type').value = type;
		var frmData = new FormData(form);

		$.ajax({
			enctype: 'multipart/form-data',
			type : 'post',           // 타입 (get, post, put 등등)
			url : '/excel/getDiameterForTreeName.dor',           // 요청할 서버url
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