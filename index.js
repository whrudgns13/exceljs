let btn = document.querySelector("#btn");
btn.addEventListener("click", apExcelDownload);

function apExcelDownload() {
  let aProductCode;
  let aProductComponent;
  
  jQuery.ajax({
    url: "./productCode.json",
    type: "GET",
    async: false,
    success: function (data) {
      aProductCode = data;
    },
    error: function (oError) {
      console.log(oError)
    }
  });

  jQuery.ajax({
    url: "./productComponent.json",
    type: "GET",
    async: false,
    success: function (data) {
      aProductComponent = data;
    },
    error: function (oError) {
      console.log(oError)
    }
  });

  //워크북 생성
  let wb = new ExcelJS.Workbook();
  
  //시트생성
  let ws = wb.addWorksheet("test");
  
  addProductRow({ws,aProductCode,aProductComponent});
  columnAlignWidth(ws);

  function addProductRow({ws,aProductCode,aProductComponent}){
    if (aProductCode.length) {
      //카테고리 행 삽입
      aProductCode.forEach(productCode => {
        let aCode = ["제품코드", "제품구성", "상품유형"];
        let aCodeValue = [productCode.MATNR, productCode.MATTPX, productCode.MATKDX];
  
        ws.addRow(aCode).eachCell( cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '000099' }
          };
  
          cell.font = {
            color: { argb: 'fffffff' }
          };
        })
  
        ws.addRow(aCodeValue);
        
        let aComponent = aProductComponent.filter(productComponent => {
          return productCode.POSNR === productComponent.POSNR;
        });
  
        //구성품 삽입
        if (aComponent.length) {
          let aComponentHeader = ["신규여부", "구성품코드", "구성품명"];
          ws.addRow(aComponentHeader).eachCell( cell => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '6666ff' }
            };
  
            cell.font = {
              color: { argb: 'fffffff' }
            };
          });
  
          aComponent.forEach( component => {
            ws.addRow([component.OMTTPX, component.OMTNR, component.OMTNRX]).outlineLevel = 1;
          })
  
        }  
      })
    }
  }
  
  //셀 중간정렬 셀 크기
  function columnAlignWidth(ws){
    ws.columns.forEach( column => {
      let maxLength = 0;
      column.eachCell( cell => {
        let columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = maxLength < 10 ? 10 : maxLength+10;
      column.alignment =  { vertical: 'middle', horizontal: 'center' };
    });
  }
  
  //워크북 내보내기
  wb.xlsx.writeBuffer().then( buffer => {
    saveAs(
      new Blob([buffer], { type: "application/octet-stream" }),"ap구성품.xlsx"
    );
  });

}

