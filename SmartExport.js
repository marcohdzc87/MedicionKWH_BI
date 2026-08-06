var laySet = [];
define( ["jquery",
		 "qlik",
		 "css!./css/SmartExport",
		 "./js/FileSaver",
		 "./Properties"
		 ],
	
	function ($,qlik,cssContent,FileSaver,properties) {
		'use strict';	
		
		function toggleId (currentSelections) {	
			var vWidth = '20000px';
			var vHeight = '50000px';
			var vTableType = '';

		    $( '.qv-object, .qv-panel-sheet' ).each( function ( i, el ) {
				var s = angular.element( el ).scope();

				if ( s.layout || (s.$$childHead && s.$$childHead.layout) ) {
					if(s.model.layout.qInfo.qType == 'table' || s.model.layout.qInfo.qType == 'pivot-table'){
						var layout = s.layout || s.$$childHead.layout, model = s.model || s.$$childHead.model;
						
						$( el ).find('.SmartExport-tooltip').remove();

						$( el ).append( 
							'<div class="SmartExport-tooltip" style="position: absolute !important; top: 8px !important; left: 8px !important; bottom: auto !important; right: auto !important; z-index: 10 !important;">' +
								'<a id="SmartExportBtn" class="SmartExport-btn" style="color:' + laySet.color + ' !important; background:' + laySet.background + ' !important; display: flex !important; align-items: center !important; justify-content: center !important; width: auto !important; height: 34px !important; border-radius: 6px !important; cursor: pointer !important; text-decoration: none !important; padding: 0 10px !important;" title="Descargar Excel""' + s.model.layout.qInfo.qType + '">' +
									'<i class="lui-icon lui-icon--export" style="font-size: 16px !important; margin-right: 5px !important;"></i>' +
									'<span style="font-size: 12px !important; font-weight: bold !important;">Exportar</span>' + 
								'</a>' +							
							'</div>' 
						);
					}
					
					$( el ).off('click', '#SmartExportBtn').on( 'click', '#SmartExportBtn', function (e) {
						e.preventDefault();
						e.stopPropagation();
						
						vTableType = this.title;
						
						model.getProperties().then( function ( reply ) {
							var app = qlik.currApp();
							var vObjectId = reply.qInfo.qId;

							app.getObject('CurrentSelections').then(function(model){
								
								var vModal = '<div id="myModal" class="modal">' +
								'<div class="modal-content">' +
								  '<span class="close" style="position: absolute; top: 12px; right: 16px; font-size: 22px; font-weight: bold; cursor: pointer; color: #888; z-index: 1002;">&times;</span>' +
								  '<div class="modal-body" style="position: relative; overflow: hidden;">' +
									'<div id="ExportEditor">' +
										'<button id="XLSButton" class="XLSbtn-modern" disabled>' +
											'<div id="BtnSpinner" class="modern-spinner"></div>' +
											'<span id="BtnText">Cargando datos...</span>' +
										'</button>' +
									'</div>' +
									'<div id="QVSmartExport02" style="position: absolute; top: -99999px; left: -99999px; width: ' + vWidth + '; height: ' + vHeight + ';"></div>' +
								  '</div>' +
								'</div>' +
								'</div>';									
								
								if(document.getElementById('myModal')){										
									document.getElementById('myModal').remove();
								}

								$( document.body ).append( vModal );
							
								var modal = document.getElementById('myModal');
								var span = document.getElementsByClassName("close")[0];
								var XLSButton = document.getElementById('XLSButton');
								
								modal.style.display = "block";
							
								app.getObject('QVSmartExport02', vObjectId).then(function(visModel) {
									// FASE 1: espera fija de 3s para que Qlik inicialice el objeto
									// y empiece a renderizar el DOM antes de sondearlo.
									setTimeout(function() {
										// FASE 2: sondeo cada 100ms hasta detectar <td> en el DOM
										// (indica que Qlik terminó de pintar los datos).
										// Máximo 30 segundos adicionales antes de mostrar error.
										var vPollMax      = 300; // 300 × 100ms = 30s
										var vPollCount    = 0;
										var vPollInterval = setInterval(function() {
											vPollCount++;
											var container = document.getElementById('QVSmartExport02');
											var tds       = container ? container.getElementsByTagName('td') : [];
											var listo     = tds.length > 0;

											if(listo || vPollCount >= vPollMax) {
												clearInterval(vPollInterval);

												var XLSButton = document.getElementById('XLSButton');
												var btnSpinner = document.getElementById('BtnSpinner');
												var btnText    = document.getElementById('BtnText');

												if(XLSButton) {
													if(listo) {
														XLSButton.disabled = false;
														if(btnText) btnText.innerText = "Descargar Excel";
														if(btnSpinner) {
															btnSpinner.className          = "lui-icon lui-icon--download";
															btnSpinner.style.animation    = "none";
															btnSpinner.style.border       = "none";
															btnSpinner.style.borderRadius = "0";
															btnSpinner.style.width        = "auto";
															btnSpinner.style.height       = "auto";
															btnSpinner.style.color        = "#fff";
														}
													} else {
														// 33s totales sin datos: mostrar error
														if(btnText) btnText.innerText = "Error: sin datos, intente de nuevo";
														if(btnSpinner) {
															btnSpinner.className          = "lui-icon lui-icon--warning";
															btnSpinner.style.animation    = "none";
															btnSpinner.style.border       = "none";
															btnSpinner.style.borderRadius = "0";
															btnSpinner.style.width        = "auto";
															btnSpinner.style.height       = "auto";
															btnSpinner.style.color        = "#e53e3e";
														}
													}
												}
											}
										}, 100);
									}, 5000); // 3 segundos de espera inicial fija
								});

								XLSButton.onclick = function() {
									if(XLSButton.disabled) return;

									var elements = document.getElementById('QVSmartExport02').getElementsByClassName('lui-button');
									var elementHeader = document.getElementById('QVSmartExport02').getElementsByClassName('hidden-screen-reader-label');
									
									if(elementHeader.length > 0) {
										elementHeader[0].remove();
									}

								    while(elements.length > 0){
								    	elements[0].parentNode.removeChild(elements[0]);									    	
								    }
								    
								    var vTextSelections = '';
								    if(laySet.selections){
										var iterator = currentSelections.length;
										vTextSelections = '<i><u><b style="color:#1f7044">Selecciones</b></u><br>';											
										
										if (iterator == 0) {
											vTextSelections += 'none</i>';
										}
										
										for (var ai = 0;ai < iterator;ai++ ) {
										    var value = currentSelections[ai];
											if (value.qSelectedCount > 6) {
											    vTextSelections += '<a style = "color:#375623">' + value.qField + ' : ' + value.translation + '</a><br>';
											} else {
											    vTextSelections += '<a style = "color:#375623">' + value.qField + ' : ' + value.qSelected + '</a><br>';
											}
										}
										vTextSelections += '</i>';
									}
									
									modal.style.display = "none";   

									var vEncodeHead = '<html><head><meta charset="UTF-8"></head>';
									var vEncodeCode = document.getElementById('QVSmartExport02');	
									
									var labels = vEncodeCode.getElementsByTagName("label");
									if(labels.length > 1 && labels[1].title == ""){
										labels[1].remove();	
									}
									
									for(var vLabel = 0;vLabel < labels.length;vLabel++){
										if(labels[vLabel].innerText == 'Table' || labels[vLabel].innerText == 'Pivot table' || labels[vLabel].innerText == 'Load previous' || labels[vLabel].innerText == 'Load more'){
											labels[vLabel].remove();
										}
									}

									var header = vEncodeCode.getElementsByTagName("header");
								    
									if(vTableType == 'pivot-table'){
								    	var tdelements = vEncodeCode.getElementsByTagName('col');
								    	for(var vtd = 0;vtd < tdelements.length;vtd++){								    		
								    		tdelements[vtd].style.cssText = 'width:135px';								    		
								    	}

										// --- LÓGICA DE DATOS: RELLENAR HACIA ABAJO (DESCOMBINAR ROWSPAN) ---
										// ESTA LÓGICA SOLO AFECTA LAS CELDAS DE DATOS (TD). NO TOCA LOS TH (ENCABEZADOS).
										var trElements = vEncodeCode.getElementsByTagName('tr');
										var pendingClones = []; 

										for (var r = 0; r < trElements.length; r++) {
											var row = trElements[r];
											
											// Ignorar filas que sean puramente encabezado
											if (row.getElementsByTagName('th').length > 0 && row.getElementsByTagName('td').length === 0) {
												continue; 
											}

											var cells = Array.prototype.slice.call(row.children);
											var physicalIndex = 0;
											var virtualIndex = 0;

											while (virtualIndex < pendingClones.length || physicalIndex < cells.length) {
												if (pendingClones[virtualIndex]) {
													var cloneData = pendingClones[virtualIndex];
													var newCell = cloneData.element.cloneNode(true);
													
													newCell.removeAttribute('rowspan'); 
													newCell.style.width = "135px";
													
													var cSpan = parseInt(newCell.getAttribute('colspan') || '1');
													
													if (physicalIndex < row.children.length) {
														row.insertBefore(newCell, row.children[physicalIndex]);
													} else {
														row.appendChild(newCell);
													}
													
													cloneData.remaining--;
													if (cloneData.remaining <= 0) {
														for(var cs=0; cs<cSpan; cs++){
															pendingClones[virtualIndex + cs] = null;
														}
													}
													
													virtualIndex += cSpan;
												} else if (physicalIndex < cells.length) {
													var cell = cells[physicalIndex];
													var rSpan = parseInt(cell.getAttribute('rowspan') || '1');
													var cSpan = parseInt(cell.getAttribute('colspan') || '1');

													cell.removeAttribute('rowspan');
													cell.style.width = "135px";

													if (rSpan > 1) {
														for(var cs=0; cs<cSpan; cs++) {
															pendingClones[virtualIndex + cs] = {
																element: cell,
																remaining: rSpan - 1
															};
														}
													}
													virtualIndex += cSpan;
													physicalIndex++;
												} else {
													break;
												}
											}
										}
										// --- SE ELIMINÓ TODO EL CÓDIGO QUE INTENTABA ALTERAR LOS ENCABEZADOS (TH) ---

								    } else {
										var tdelements = vEncodeCode.getElementsByTagName('td');
										for(var vtd = 0; vtd < tdelements.length; vtd++){
											var styles = tdelements[vtd].style.cssText;
											tdelements[vtd].style.cssText = styles + ';width:135px;';
										}
										var thelements = vEncodeCode.getElementsByTagName('th');
										for(var vth = 0; vth < thelements.length; vth++){
											var stylesth = thelements[vth].style.cssText;
											thelements[vth].style.cssText = stylesth + ';width:135px;';
										}
									}
									
									if(!laySet.title){											
										header[0].remove();
									}else{
										var H1s = header[0].getElementsByTagName("h1");
										if(H1s.length > 0){
									        H1s[0].outerHTML = H1s[0].outerHTML.replace("<h1", "<div").replace("</h1>","</div>");									        									        	
											if(!laySet.subtitle && reply.subtitle != ""){
												var header = vEncodeCode.getElementsByTagName("header");
												var subtitle = header[0].getElementsByTagName("h2");
												subtitle[0].remove();
											}
										}
									}

									if(!laySet.footer){
										var footer = vEncodeCode.getElementsByClassName("qv-footer-wrapper");
										if(footer.length > 0) footer[0].remove();
									}
									
									var vEncodeBody = vEncodeCode.innerHTML.replace("Load previous", "").replace("Load more", "");
									var blob = new Blob([vEncodeHead + vEncodeBody + vTextSelections + '</html>'], {
										type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8"
										});
									
									var vFileName = reply.title || vObjectId;
									var vNow = new Date();
									var vFecha = vNow.getFullYear() + '-' + 
									             String(vNow.getMonth()+1).padStart(2,'0') + '-' + 
									             String(vNow.getDate()).padStart(2,'0') + '_' + 
									             String(vNow.getHours()).padStart(2,'0') + 
									             String(vNow.getMinutes()).padStart(2,'0');
									saveAs(blob, vFileName + '_' + vFecha + '.xls');
								};

								span.onclick = function() {
									modal.style.display = "none";    									    									
								};									
							})								
						})
					})
				} 
			})			
		}

		return {
			initialProperties: {
				version: 1.0,
				showTitles: false
			}, 
			definition : properties,
			paint: function ( $element,layout ) {	
				var app = qlik.currApp();
				app.getList("CurrentSelections", function(reply) {
					var mySelectedFields = reply.qSelectionObject.qSelections;					
					laySet = { "title":layout.titlebool,"subtitle":layout.subtitlebool,"footer":layout.footerbool,"selections":layout.selectionsbool,"background":layout.iconbackground.color,"color":layout.iconcolor.color};
					
					toggleId(mySelectedFields);
				})
			}
		};
	});