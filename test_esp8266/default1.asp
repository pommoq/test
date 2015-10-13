<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<title>Document</title>
	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css">
	<script src="https://code.jquery.com/jquery-2.1.4.min.js"></script>
	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js"></script>

	<link href="https://gitcdn.github.io/bootstrap-toggle/2.2.0/css/bootstrap-toggle.min.css" rel="stylesheet">
	<style>	.toggle-handle.btn-lg {width: 250px;}</style>
	<script type="text/javascript">
	var host = "./";
		$(document).ready(function() {
			$('.sw').prop({"data-width":"500","data-toggle":"toggle","data-size":"large"})	

			setState('s1',"p1");
			setState('s2',"p2");
			setState('s3',"p3");
			setState('s4',"p4");

			$('.sw').change(function() {
				console.log( this.id +': ' + $(this).prop('checked'))

				//sendCommand(this.id,this,$(this).prop('checked'))
				//aa();
			})
			$('.form-group').click(function() {
				console.log( this.id +'-click: ' + $(this)[0].children[0].id)
				//aa();
				obj = $(this).find(".sw")[0]
				sendCommand(obj.id,obj,!($(obj).prop('checked')))

			})

		})
		function aa(){
			alert(0)
		}
		function sendCommand(command,obj,state){
			var geturl = host + command;
			if(state){
				geturl += "1";
			}else{
				geturl += "0";
			}
			$.get(geturl, function(data) {
				
				if(data.substring(2,1)=='1'){
					//$('#'+obj.id).prop('checked', true).change()
					$('#'+obj.id).bootstrapToggle('on')
				}else{
					$('#'+obj.id).prop('checked', false).change()
				}
				
			});
		}		
		function setState(command,objid){
			
			$.get(host + command, function(data) {
				//alert(data)
				$('#'+objid).prop('disabled',false);
				if(data.substring(2,1)=='1'){
					$('#'+objid).bootstrapToggle('on')
				}else{
					$('#'+objid).bootstrapToggle('off')
				}
			});
		}
	</script>
</head>
	
	<script src="https://gitcdn.github.io/bootstrap-toggle/2.2.0/js/bootstrap-toggle.min.js"></script>
<body>
<nav class="navbar navbar-default" role="navigation">
	<!-- Brand and toggle get grouped for better mobile display -->
	<div class="navbar-header">
		<button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-ex1-collapse">
			<span class="sr-only">Toggle navigation</span>
			<span class="icon-bar"></span>
			<span class="icon-bar"></span>
			<span class="icon-bar"></span>
		</button>
		<a class="navbar-brand" href="#">Smart Switch</a>
	</div>

	<!-- Collect the nav links, forms, and other content for toggling -->
	<div class="collapse navbar-collapse navbar-ex1-collapse">
		<ul class="nav navbar-nav navbar-right">
			<li><a href="#">Action</a></li>
		</ul>
	</div><!-- /.navbar-collapse -->
</nav>
<div class="container" align="center">
	<form>
  	<div class="form-group">
			<input type="checkbox" name="p1" id="p1" data-toggle="toggle" disabled="disabled"  data-width="100%" data-height="75" class="sw">	
	</div>
	<div class="form-group">
			<input type="checkbox" name="p2" id="p2"  disabled="disabled" class="sw">	
	</div>
	<div class="form-group">
			<input type="checkbox" name="p3" id="p3" data-toggle="toggle" disabled="disabled" data-size="large" class="sw">
	</div>
	<div class="form-group">
			<input type="checkbox" name="p4" id="p4" data-toggle="toggle" disabled="disabled" data-size="large" class="sw">
	</div>
	</form>
</div>

</body>
</html>