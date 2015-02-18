//if ($('input[name=topic]').val()=='11098'){
//    $('#content_section').css({'background':'none','background-color':'red'})
//}
function SMF_ProcessSensitives(){
    var Str=document.body.innerHTML;
    var BBcode='Sensitive';
    var Pstart=Str.indexOf('['+BBcode+']');
    if (Pstart!=-1){
        var Pend=Str.indexOf('[/'+BBcode+']');
        var Instruction=Str.substring( Pstart + BBcode.length+2 , Pend );
        var Command=Instruction.split(':')[0];
        console.log( Command );
        if (Command=='style'){
            var style=Instruction.split(':')[1];
            var id=Instruction.split(':')[2];
            var value=Instruction.split(':')[3];
            var C=document.getElementById(id);
            C.style.background='none';
            console.log( style );
            C.style[style]=value;
        }
	}
}
SMF_ProcessSensitives();
