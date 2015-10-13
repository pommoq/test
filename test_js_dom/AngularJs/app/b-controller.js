app.controller("BController",function($scope){
	$scope.Title = "Body";
	$scope.orderFieldName = "id";
	$scope.flag_ASC = false;

	$scope.EditData = {name:"",phone:""};
	$scope.curIndex = -1;

	$scope.Tables = [];
	$scope.start = function(){
		$scope.Tables.push({id:1,name:"AABBCC",phone:"086-9922334455"});
		$scope.Tables.push({id:2,name:"CCDDEE",phone:"086-9977665544"});
		// $scope.Tables.push({id:3,name:"DDRRTT",phone:"086-0000000000"});
		// $scope.Tables.push({id:4,name:"GGBBHH",phone:"086-8877665533"});
		// $scope.Tables.push({id:5,name:"VVKKSS",phone:"086-1111111111"});
		// $scope.Tables.push({id:6,name:"PPLLOO",phone:"086-3333333333"});
		// $scope.Tables.push({id:7,name:"SSXXMM",phone:"086-2222222222"});
		// $scope.Tables.push({id:8,name:"pompom",phone:"086-1111111111"});
		// $scope.Tables.push({id:9,name:"ppppom",phone:"086-5555555555"});
		$scope.Tables.push({id:10,name:"ป้อมเองครับ",phone:"086-8888888888"});
	}

	$scope.newData = function(){
		$scope.curIndex = -1;

		$("#div-datatable").hide();
		$("#div-editbox").show();
	}

	$scope.editData = function(data){
		$scope.curIndex = $scope.Tables.indexOf(data)
		
		$scope.EditData.name = data.name;
		$scope.EditData.phone = data.phone;

		$("#div-datatable").hide();
		$("#div-editbox").show();


	}

	$scope.saveData = function(){
		if($scope.curIndex < 0){
			$scope.Tables.push({name: $scope.EditData.name , phone : $scope.EditData.phone})
		}else{
			$scope.Tables[$scope.curIndex].name = $scope.EditData.name
			$scope.Tables[$scope.curIndex].phone = $scope.EditData.phone
		}
		
		

		$scope.EditData = {name:"",phone:""};
		$("#div-datatable").show();
		$("#div-editbox").hide();
	}

	$scope.cancelData = function(){
		$("#div-datatable").show();
		$("#div-editbox").hide();

		$scope.EditData = {name:"",phone:""};
	}

});