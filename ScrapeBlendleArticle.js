//scrape Blendle article, copy to clipboard, browse back

var H='<div style="height:800px;overflow:scroll;background-color:white;font-size:20px;z-index:999999;position:absolute;margin:40px;padding:20px;" onclick="$(this).remove();">';
$('.item-paragraph').each( function(index,element){ H+='<p>'+$(element).text()+'</p>' } );
H+='</div>';
$('body').append( H );
window.clipboardData.setData( 'Text' , H );
window.history.back();
