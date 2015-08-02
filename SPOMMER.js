/*global SP,console,DOMParser,$,window,document */
/*
* SPOMinspector
*
* Copyright (c) 2014 Danny Engelman   www.ViewMaster365.com  All rights reserved.
*
* Permission is hereby granted, free of charge, to any person obtaining a
* copy of this JavaScript library and associated documentation files (the "Software"),
* to deal in the Software without restriction, including without limitation
* the rights to use, copy, modify, merge, publish, distribute, sublicense,
* and/or sell copies of the Software, and to permit persons to whom the
* Software is furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in
* all copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
* FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
* DEALINGS IN THE SOFTWARE.
*
*/
function SPOMinspector( _properties ){
/*
	SharePoint Object Model Inspector for Apps
	
	Calls the HostWeb from a SharePoint App with async/recursive REST calls
	
	Returns a Promise and the read Webs/Lists/Fields as a _SPOM object
	
	options:
		- cleanresponse - true deletes all __metadata and __deferred references from the response Object
		- Properties: array of named deferred Properties to (recursive) load
			fails silent for unexisting properties
			SharePoint Property value is replaced by an Array of read results.
			e.g. Fields : Array[69]
							[0] - (Field)Object
							[1] - (Field)Object
		- Callback(s) - functions to execute after iets loaded REST response
			e.g. use to update UI progress to the users
			fails silent for unexisting functions

	returned _SPOM object contains all object with GUID as the Key
		+ List : GUID array of all found Lists
		+ Web
		+	___RESTcallcount : (progressive) number of REST calls made

		each read Object gets additional data:
			___SPType : List/Field/View extracted from __metadata.type
			___listID : GUID pointing back to List this field is in
		List Objects:
			___defaultViewUrl : GUID
		
Example code:		
		var hostWeb = new SPOMinspector( { RESTendpoint:['/Web']
											, cleanresponse:false
											, Types:{
														Callback:false
														,Web:{ Properties:['RegionalSettings','Lists'] , Callback:false }
														,List:{ Properties:['Fields','Views'] , Callback:justloadedList }
														,View:{ Properties:[] , Callback:false}
														,Field:{ Properties:[] , Callback:false}
													}
											, ReadyCallBack:createViewMasterSPliststructure });
		hostWeb.done(function(_SPOM){
			//your code here
		});
	*/
	_properties = _properties || { RESTendpoint:'/Web/Lists', 
											 cleanresponse:false,
											//filterLists:['Tasks','Documents','Second TaskList'],
											Types:{
														Callback:false,
														Web:{ Properties:['Webs']	, Callback:false },
														List:{ Properties:['Fields','Views'] , Callback:false },
														View:{ Properties:[] , Callback:false},
														Field:{ Properties:[] , Callback:false}
													},
											 ReadyCallBack:false };
	var _RESTcallcount=0;//total number of REST calls done
	var _traceREST=_properties.traceREST || true;//true outputs to console
    var _SPOMdef=new $.Deferred();//return a promise which resolves after the first run
    var _startUrl=_properties.RESTendpoint || '/';
    var _traces=[];//array of strings filled by trace calls
	var _SPOM={//object to be filled by SPOMinspector
				List:[],
				Web:[],
				Webs:{}
				};

	var _DOMparser = new DOMParser();//used to convert SchemaXml attributes to Object

    var _trace=function(type,p1,p2,p3,p4,p5,p6){
		var none='';
		var line=p1||none+' , '+p2||none+' , '+p3||none+' , '+p4||none+' , '+p5||none+' , '+p6||none;
		_traces.push( line );
		if (window.console){
			if (typeof p1=='boolean'){
				if(p1){ console[type](p2||none,p3||none,p4||none,p5||none,p6||none); }
			} else {
				console[type](p1||none,p2||none,p3||none,p4||none,p5||none,p6||none);
			}
		}
    };
    var consoleinfo=function(p1,p2,p3,p4,p5,p6){
        _trace('info',p1,p2,p3,p4,p5,p6);
    };
    var consoleerror=function(p1,p2,p3,p4,p5,p6){
        _trace('error',p1,p2,p3,p4,p5,p6);
    };
	var getQueryStringParameter=function (paramToRetrieve) {
		var i;
		var	params = document.URL.split("?")[1];
		if(params){
			params=params.split("&");
			for (i = 0; i < params.length; i = i + 1) {
				var singleParam = params[i].split("=");
				if (singleParam[0] == paramToRetrieve) {
					return singleParam[1];
				}
			}
		}
	};
    var _extractGUID=function(S,_GUIDstr){//extract guide from strings: guid('xxxxx')
        var _GUIDmarker = _GUIDstr+"(guid'";
        var _GUIDend = "')";
		var _GUID = false;
        if(S.indexOf(_GUIDmarker)>-1){
            _GUID=S.split(_GUIDmarker)[1].split(_GUIDend)[0];
        }
		//todo: add extra code to extract from: {xxxxx}
        return( _GUID );
    };
    this.loadAll = function(_startUrl) {
        var _dfds = []; // deferreds for the current level.
        var _urls = []; // urls for current level.
			_startUrl.forEach(function(_url){
				_urls.push(_url);
			});
        var _responses = []; // _responses for current level.
        var _cache = {}; // object to map urls to promises. for future use
		
		//called for every REST response, converts the data object in my preferred format as ___ViewMaster object
		//changes _deferred values of Views, Fields, Choices to a proper Array of values
        var _parseSPRESTresponse = function (data) {// given the responseText, add any referenced urls to the urls array
            var _SPtype=data.__metadata.type.split('.')[1];
			var _WebId=_SPOM.Webs[data.url];
            var _listGUID=_extractGUID(data.__metadata.id,"Lists");
            data.___ViewMaster={
                SPType:_SPtype,
                listID:_listGUID,
                url:data.url,
                RESTcall:data.__metadata.uri
            };
			if (_properties.cleanresponse){//if boolean then clean up the data object (to make it smaller)
				delete( data.__metadata );
					var key;
					for(key in data){
						if(data[key] && data[key].hasOwnProperty('__deferred')){
							data[key]='__deferred';
							//data[key]=data[key].__deferred.uri;
						}
					}
			}
//            if(_traceREST){consoleinfo('response',_SPtype,data.TypeAsString,_listGUID,data.url,data);}

            if(_SPtype==='RegionalSettings'){
				_SPOM[ _WebId ][_SPtype]=data;
            }
			//for all Fields found
            if(_SPtype.indexOf('Field')===0){

				//reformat the Choices
                if(data.hasOwnProperty('Choices')){
                        data.___ViewMaster.Choices=data.Choices.results;
                }

				//create Url to Edit Column settings
				data.___ViewMaster.editUrl =data.url;
				data.___ViewMaster.editUrl+="/_layouts/15/FldEdit.aspx?List=%7B"+_listGUID+"%7D&Field="+data.InternalName;
				
				//store the SchemaXml string into an Object 
				//to get DisplayName, because get_title() does NOT return the (localized) Display Name when run from App
				var n,SchemaXml = _DOMparser.parseFromString(data.SchemaXml, "application/xml").getElementsByTagName('Field')[0];
				if(SchemaXml && SchemaXml.attributes){
					data.___ViewMaster.SchemaXml={};
					for(n=0;n<SchemaXml.attributes.length;n++){
						var _attr=SchemaXml.attributes[n];
						data.___ViewMaster.SchemaXml[ _attr.name ]=_attr.value;
					}
				}

				//store the data in 2 locations
                _SPOM[ _listGUID ].Fields.push(data);
				_SPOM[ _listGUID ].___ViewMaster.HasFields[data.InternalName]=data;
				
				//All the possible Field types
                switch(_SPtype){
                    case('Field'):
                    case('FieldComputed'):
                    case('FieldNumber'):
                    case('FieldDateTime'):
                    case('FieldUser'):
                    case('FieldMultiLineText'):
                    case('FieldGuid'):
                    case('FieldText'):
                    case('FieldUrl'):
                        break;
                    case('FieldCalculated'):
                        break;
                    case('FieldLookup'):
                        break;
                    case('FieldChoice'):
                    console.log(data.Title,data.Choices.results);
                        break;
                    case('FieldMultiChoice'):
                        break;
                    default:
						//We should never come here
                        //consoleinfo('default',data,_SPOM[ _listGUID ].Title,data.Title);
                }
                _SPtype='Field';
            }
			//Process View definitions
            if(_SPtype=='View'){
				//link to edit view page
				data.___ViewMaster.editUrl=data.url+"/_layouts/15/ViewEdit.aspx?List=%7B"+_listGUID+"%7D&View=%7B"+data.Id+"%7D";
			
                _SPOM[ _listGUID ].Views.push(data);
                if(data.DefaultView){
                    _SPOM[ _listGUID ].___ViewMaster.defaultViewId=data.Id;
                    _SPOM[ _listGUID ].___ViewMaster.defaultViewUrl=data.ServerRelativeUrl;
                    _SPOM[ _listGUID ].___ViewMaster.editdefaultViewUrl=data.___ViewMaster.editUrl;
                }
            }
            if(_SPtype=='List'){
				if( _listGUID===false ) { 
					if(_traceREST){consoleerror( 'Missing listGUID' , data.Title );}
				} else {
					var _filterLists=_properties.filterLists;//false; used for testing to filter out list by List.Title
					if( !_filterLists || (_filterLists && _filterLists.indexOf(data.Title)>-1)){
						_SPOM[ _SPtype ].push(data.Id);
						var _GUID=data.Id;
						data.___ViewMaster.HasFields={};
						data.___ViewMaster.editUrl=data.url+"/_layouts/15/listedit.aspx?List=%7B"+_listGUID+"%7D";
						_SPOM[ _GUID ]=data;
					//	_SPOM[ _WebId ].Lists.push(data);//do NOTdouble the whole structure
						_SPOM[ _WebId ].Lists.push(_GUID);//refer to GUID in root
						//Call the REST endpoint for Fields and Views of this List
						//in array for future expandability
						_properties.Types[_SPtype].Properties.forEach(function(property){
							try{
								var newRESTendpoint="/Web/Lists(guid'"+_GUID+"')/"+property;
								if(_traceREST){consoleinfo(newRESTendpoint);}
								_urls.push(newRESTendpoint);//add this REST endpoint to the queue
								_SPOM[ _listGUID ][property]=[];//reset Fields or Views object to Array, for storing results
							}
							catch(e){ consoleerror(e);}
						});
					}
				}
            }
            if(_SPtype==='Web'){
				_SPOM[ _SPtype ].push(data.Id);
				_SPOM.Webs[data.url]=data.Id;
				data.___ViewMaster.editUrl=data.url+"/_layouts/15/settings.aspx";
				_SPOM[ data.Id ]=data;
				//Call the REST endpoint for Fields and Views of this Web
				//in array for future expandability
				_properties.Types[_SPtype].Properties.forEach(function(property){
					try{
//            console.log(property,data[property].hasOwnProperty('__deferred'),data[property].__deferred.uri);
						if(data[property].hasOwnProperty('__deferred')){
//							console.log('do',data[property].__deferred.uri);
							var newRESTendpoint=data[property].__deferred.uri;
							_SPOM[ data.Id ][ property ]=[];
							if(_traceREST){consoleinfo(newRESTendpoint);}
							_urls.push(newRESTendpoint);//add this REST endpoint to the queue
						}
					}
					catch(e){ consoleerror(e.message,data);}
				});
            }
            //check if there is a Callback function for this Type            //or a generic Callback for Types
            try{
                var _CallbackParse= _properties.Types.Callback;
                if(_CallbackParse){//if so pass the data object to the synchronous callback function
					var _CallBackData = data.hasOwnProperty('Title') ? data.Title : data ;
                    _CallbackParse( _CallBackData );
                }
            }
            catch(e){
				consoleerror(e);
			}
        };
		//execute one REST call
        var _callRESTurl = function(url) {   
            var _dfd;
            if(_cache.hasOwnProperty(url)){// use _cached promise (FOR FUTURE USE)
                _dfd = _cache[url];// if it is already resolved, any callback attached will be called immediately.
                _dfds.push(_cache[url]);
            } else {
                _dfd = $.Deferred();
                var _url,_executor,isAPP=false;
				_RESTcallcount++;
                if(isAPP){
					var _RESTheaders={ "Accept": "application/json; odata=verbose"};
					//_spPageContextInfo.webAbsoluteUrl
					var _hostweburl = decodeURIComponent(getQueryStringParameter('SPHostUrl'));
					var _appweburl = decodeURIComponent(getQueryStringParameter('SPAppWebUrl'));
					_executor = new SP.RequestExecutor( _appweburl );
					_SPOM.hostWeb = _hostweburl;
					url=url.replace(_hostweburl+'/_api','');//remove _api part from full (deferred) uris
					_SPOM.appWeb = _appweburl;
						_url=_appweburl;
						_url+="/_api/SP.AppContextSite(@target)"+url;
						_url+="?@target='" + _hostweburl + "'";            
					if(_traceREST){consoleinfo('Calling:',_RESTcallcount,_url);}
					_executor.executeAsync(
						{url: _url, type: 'GET', headers: _RESTheaders,
							success:function(_data){
										if(_data.hasOwnProperty('body')){//handle different type of responses
											_data = JSON.parse( _data.body );
										}
//TESTING UPDATING DOM FOR PROGRESS, SEEMS CHROME BLOCKS
//setTimeout( function(){
//	var SPOMMERstatus=document.getElementById("SPOMMERstatus");
//	if(SPOMMERstatus && response.Title){SPOMMERstatus.innerHTML=response.Title;}
//}, 0);										
										_data.url=_hostweburl;//_add the url because the info is only in this 'this' object;
										_dfd.resolve(_data);// resolve and pass response.
									},
							
							error:function (_response,_errorCode,_errorMessage){
										consoleerror(_errorCode,_errorMessage);
										_dfd.resolve(null);// resolve and pass null, so this error is ignored.
									}
						});
                } else {//REST calls directy on current SP site
                _hostweburl=_spPageContextInfo.webAbsoluteUrl;
                if(url.indexOf('_api')===-1){
					_url=_hostweburl+'/_api'+url;//'/_api/Web/';
                } else {
                	_url=url;
                }
					REST={
						url:_url,
						type:'GET',
						data:'',
						success:function(_data){
										if(_data.hasOwnProperty('body')){//handle different type of responses
											_data = JSON.parse( _data.body );
										}
//										console.info('Received',_url,_data);
										_data.url=_hostweburl;//_add the url because the info is only in this 'this' object;
										_dfd.resolve(_data);// resolve and pass response.
									},
							
							error:function (_response,_errorCode,_errorMessage){
										consoleerror(_errorCode,_errorMessage);
										_dfd.resolve(null);// resolve and pass null, so this error is ignored.
									}
					}
					if(_traceREST){consoleinfo('Calling:',_RESTcallcount,url,_url);}
						$.ajax({
						  url: REST.url,
						  type: REST.type,
						  data: REST.data,
						  headers: { 
							"X-HTTP-Method":REST.httpmethod,
							"X-RequestDigest": $("#__REQUESTDIGEST").val(),
							"accept": "application/json;odata=verbose",
							"content-type": "application/json;odata=verbose"
							//"content-length": <length of body data>
						  },
						  success: REST.success,
						  error: REST.error
						});
                }
				_dfds.push(_dfd.promise());//store promisses
				_cache[url] = _dfd.promise();
            }
            var _addResponse = function( _data , _response ){
                _response.url = _data.url;
                _responses.push( _response );
            };
            _dfd.done(function(_data) {// when the request is done, add response to array.
                if(_data && _data.d){
                    if(_data.d.hasOwnProperty('results')){//multiple _responses
                        _data.d.results.forEach(function(_response){
                            _addResponse( _data , _response );
                        });
                    } else {
                        _addResponse( _data , _data.d );//single response
                    }
					//if(_traceREST){consoleinfo('request done',_data.d.results);}
                }
            });
        };
		//when all promisses are resolved, check for a Callback function, execute, then report back were all done
		var _SPOMinspectorReadyCheck=function( _SPOM ){
			if( _properties.ReadyCallBack || false){ 
				_properties.ReadyCallBack( _SPOM );//call callback function
			}
			_SPOMdef.resolve( _SPOM );
		};
		//executes REST call for one level (eg. List level then calls all Fields and Views)
        var _loadLevel = function () {
            _dfds = [];
            _responses = [];
            _urls.forEach(function(RESTcall){
                _callRESTurl( RESTcall );
            });
            $.when.apply($, _dfds).done(function(){//level is done loading. each done function above has been called already, 
                // so _responses array is full.
                _urls = [];
                //consoleinfo('_responses',this.headers,_responses);
                _responses.forEach(function(response){// parse all the _responses for this level.
                    _parseSPRESTresponse( response );// this will refill _urls array.
                });
                if(_urls.length === 0) {//all done
					_SPOM._RESTcallcount=_RESTcallcount;
					_SPOMinspectorReadyCheck( _SPOM );
                } else {
                    _loadLevel();//recursive call
                }
            });
        };
		//do the App code or use the hardcoded Debuglist if that exists
		if (window.hasOwnProperty('SP')){
console.info(_startUrl,_properties);
			_loadLevel();
		} else {
			var S="_SPOMinspectorDebugList";
			_SPOMinspectorReadyCheck( window.hasOwnProperty(S) ? window[ S ] : false );
		}
    };//loadAll
    this.loadAll( _startUrl );
    return _SPOMdef.promise();
}//SPOMinspector
	function getSPOM(){
		var SPOMprogress=function(_data){
			//console.log('progress',_data);
		};
		var hostWeb = new SPOMinspector( { RESTendpoint:['/Web']
											,traceREST:true
											, cleanresponse:false
											, filterLists:false//['VM Taken','Tasks','Documents','Second TaskList']
											, Types:{
														Callback:SPOMprogress
														,Web:{ Properties:['RegionalSettings','Lists'] , Callback:false }
														,List:{ Properties:['Fields','Views'] , 
														Callback:false }
														,View:{ Properties:[] , Callback:false}
														,Field:{ Properties:[] , Callback:false}

													}
											//, ReadyCallBack:$scope.createViewMasterSPliststructure });
											, ReadyCallBack:false });
		hostWeb.done(function(_SPOM){
			console.info( 'SPOMconnector.done',_SPOM );
			var D=JSON.stringify(_SPOM);
			D=D.replace(/</gi,'XXXX');
			D='<pre style="word-wrap:break-word">var _SPOMinspectorDebugList ='+D+'</pre>';
			console.log( D.length );
			//document.getElementById('body').innerHTML=D;
			//document.body.innerHTML=D;
		});
	}
console.clear();
getSPOM();
