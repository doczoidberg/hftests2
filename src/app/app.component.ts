import { Component } from '@angular/core';
import { HyperFormula } from 'hyperformula';
import sockeldata from '../assets/_deltatetraneu.json';

const options = {
  licenseKey: 'agpl-v3',
  language: 'enGB'
  //  thousandSeparator: ',' as any
};


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'hftests2';
  hfInstance: any; // HyperFormula;
  constructor() {
    const hfInstance = HyperFormula.buildEmpty(options);
    this.hfInstance = hfInstance;
    // needed for excel
    hfInstance.addNamedExpression('TRUE', '=TRUE()')
    hfInstance.addNamedExpression('FALSE', '=FALSE()')
    console.log('hfInstance', hfInstance);
    var sheetnames = Object.getOwnPropertyNames((sockeldata as any).sheets);
    for (var i = 0; i < sheetnames.length; i++) {
      var name = sheetnames[i];

      var array = (sockeldata as any).sheets[name];
      const cells = array[0].map((_, colIndex) => array.map(row => row[colIndex]));
      console.log('sheet ' + name, cells);
      for (var x = 0; x < cells.length; x++) {
        for (var y = 0; y < cells[x].length; y++) {
          if (cells[x][y] === "")
            cells[x][y] = null;
        }
      }
      hfInstance.addSheet(name);
      hfInstance.setSheetContent(name, cells);
    }


    var idxRahmeneckeNord = hfInstance.getSheetId("RahmeneckeNord");
    console.log('cell', hfInstance.getCellValue({ sheet: idxRahmeneckeNord, col: 0, row: 0 })); // =IF($B$361,0,C526)


    console.log('cell', hfInstance.getCellValue({ sheet: idxRahmeneckeNord, col: 1, row: 310 })); // =RahmeneckeNord!C530
    console.log('cell', hfInstance.getCellValue({ sheet: idxRahmeneckeNord, col: 3, row: 313 })); // =IF(AND(B314=1,C318="Passbolzen"),1,0
    var idxDT = hfInstance.getSheetId("DeltaTetra");
    var ST = hfInstance.getSheetId("Statik Tetra");
    var SD = hfInstance.getSheetId("Statik Delta");
    var auswahl = hfInstance.getSheetId("Auswahl");
    var materialmapping = hfInstance.getSheetId("MaterialMapping");



    console.log('cell b2', hfInstance.getCellValue({ sheet: SD, col: 1, row: 1 })); // =IF($B$361,0,C526)
    console.log('cell b2', this.hfInstance.getCellValue(this.hfInstance.simpleCellAddressFromString("B2", SD)));
    var cellcontent = this.hfInstance.getCellValue(this.hfInstance.simpleCellAddressFromString("B2", SD)); // H169 ='Statik Tetra'!C44
    console.log("B2", cellcontent);
    this.printCellData("B2", SD);

    var ca = undefined;
    ca = hfInstance.simpleCellAddressFromString("K8", idxDT);
    hfInstance.setCellContents(hfInstance.simpleCellAddressFromString("J8", idxDT), undefined)
    console.log('J8', hfInstance.getCellValue(hfInstance.simpleCellAddressFromString("J8", idxDT)));

    console.log('K8', hfInstance.getCellFormula(ca));

    console.log('cellcaval', hfInstance.getCellValue(ca));
    ca = hfInstance.simpleCellAddressFromString("K17", idxDT); // ist leer falscherweise
    console.log('K17', hfInstance.getCellFormula(ca));
    console.log(hfInstance.getCellValue(ca));

    // lookup
    this.printCellData("I172", idxDT);

    this.printCellData("C41", SD);


    this.printCellData("F24", SD);
    this.printCellData("F30", SD);
    this.printCellData("F31", SD);
    this.printCellData("F32", SD);
    this.printCellData("F33", SD);
    this.printCellData("F34", SD);

    this.printCellData("F36", SD);
    this.printCellData("F37", SD);
    this.printCellData("F38", SD);

    // hfInstance.setCellContents(hfInstance.simpleCellAddressFromString("D34", SD), "0.72");
    this.printCellData("D34", SD);
    this.printCellData("U43", SD); // abweichung vom beispiel


    this.printCellData("B45", SD);
    this.printCellData("B46", SD);
    this.printCellData("B47", SD);
    this.printCellData("B55", SD);

    this.printCellData("R117", SD);

    this.printCellData("H168", idxDT);
    this.printCellData("H169", idxDT);
    this.printCellData("H170", idxDT);
    this.printCellData("H171", idxDT);
    this.printCellData("H172", idxDT);

    this.printCellData("I168", idxDT);
    this.printCellData("I169", idxDT);
    this.printCellData("I170", idxDT);
    this.printCellData("I171", idxDT);
    this.printCellData("I172", idxDT);

    this.printCellData("J168", idxDT);
    this.printCellData("J169", idxDT);
    this.printCellData("J170", idxDT);
    this.printCellData("J171", idxDT);
    this.printCellData("J172", idxDT);

    this.printCellData("K168", idxDT);
    this.printCellData("K169", idxDT);
    this.printCellData("K170", idxDT);
    this.printCellData("K171", idxDT);
    this.printCellData("K172", idxDT);


    ca = hfInstance.simpleCellAddressFromString("I172", idxDT);
    console.log('I172', hfInstance.getCellFormula(ca));
    console.log('cellcaval', hfInstance.getCellValue(ca));

    // hfInstance.setCellContents(hfInstance.simpleCellAddressFromString("A1", ST), "1")
    // hfInstance.setCellContents(hfInstance.simpleCellAddressFromString("A2", ST), "3")
    // hfInstance.setCellContents(hfInstance.simpleCellAddressFromString("A3", ST), "=MAX(A1:A2)")
    this.printCellData("A3", ST);

    // this.printCellData("C44", ST);
    // this.printCellData("F41", ST);
    // this.printCellData("O48", ST);
    this.printCellData("S1213", idxDT);

    //  for (var i = 1198; i < 1415; i++) {

    this.printCellData("C1198", idxDT);
    this.printCellData("D1198", idxDT);
    this.printCellData("E1198", idxDT);
    this.printCellData("F1198", idxDT);
    this.printCellData("G1198", idxDT);
    this.printCellData("H1198", idxDT);
    this.printCellData("I1198", idxDT);
    this.printCellData("J1198", idxDT);
    this.printCellData("K1198", idxDT);
    this.printCellData("L1198", idxDT);
    this.printCellData("M1198", idxDT);
    this.printCellData("N1198", idxDT);
    this.printCellData("O1198", idxDT);
    this.printCellData("P1198", idxDT);
    this.printCellData("Q1198", idxDT);
    this.printCellData("R1198", idxDT);
    this.printCellData("S1198", idxDT);


    // muss ausgeführt wreden?!
    this.printCellData("C3", materialmapping);
    this.printCellData("F5", auswahl); // ='Material-Mapping'!C3
    this.printCellData("J70", idxDT);
    this.printCellData("K70", idxDT); // =IF(ISBLANK(J70),Auswahl!$F$5,J70)
    this.printCellData("O1154", idxDT);


    //          =   "##NA="&E1154&"##GR=Hilfslinien##SG=Statikachsen##"&$E$1153&"="&E1154&"##"&$M$1153&"="&M1154&"##"&$N$1153&"="&N1154&"##"&$O$1153&"="&O1154&"##"&$P$1153&"="&P1154&"##"&$Q$1153&"="&Q1154&"##"&$R$1153&"="&R1154&"##"&$S$1153&"="&S1154&"##"&$T$1153&"="&T1154&"##"&$U$1153&"="&U1154&"##"
    // cellcontent = '="##NA="&E1154&"##GR=Hilfslinien##SG=Statikachsen##"&$E$1153&"="&E1154&"##"&$M$1153&"="&M1154&"##"&$N$1153&"="&N1154&"##"&$O$1153&"="&O1154&"##"&$P$1153&"="&P1154&"##"&$Q$1153&"="&Q1154&"##"&$R$1153&"="&R1154&"##"&$S$1153&"="&S1154&"##"&$T$1153&"="&T1154&"##"&$U$1153&"="&U1154&"##"';
    // this.hfInstance.setCellContents(this.hfInstance.simpleCellAddressFromString("T1198", idxDT), cellcontent);
    this.printCellData("T1198", idxDT); // läuft net

    this.printCellData("U1198", idxDT);
    this.printCellData("VL1198", idxDT);
    this.printCellData("W1198", idxDT);
    this.printCellData("X1198", idxDT);
    this.printCellData("Z1198", idxDT);
    this.printCellData("Z1198", idxDT);
    this.printCellData("Z1199", idxDT);
    this.printCellData("Z1200", idxDT);
    this.printCellData("Z1201", idxDT);
    this.printCellData("Z1202", idxDT);
    this.printCellData("Z1203", idxDT);
    this.printCellData("Z1204", idxDT);
    this.printCellData("Z1205", idxDT);
    this.printCellData("Z1206", idxDT);
    this.printCellData("Z1207", idxDT);
    this.printCellData("Z1208", idxDT);
    this.printCellData("Z1209", idxDT);




    this.printCellData("Z1213", idxDT);
    this.printCellData("Z1218", idxDT);
    this.printCellData("Z1231", idxDT);
    this.printCellData("Z1251", idxDT);
    this.printCellData("Z1280", idxDT);
    this.printCellData("Z1286", idxDT);
    this.printCellData("Z1412", idxDT);
    this.printCellData("Z1413", idxDT);

  }



  printCellData(cell: string, sheet) {
    var cellcontent = this.hfInstance.getCellFormula(this.hfInstance.simpleCellAddressFromString(cell, sheet)); // H169 ='Statik Tetra'!C44
    console.log(cell, cellcontent);
    var result = this.hfInstance.getCellValue(this.hfInstance.simpleCellAddressFromString(cell, sheet));
    if ((result as any)?.value === "#REF!") {
      this.hfInstance.setCellContents(this.hfInstance.simpleCellAddressFromString(cell, sheet), cellcontent);
      result = this.hfInstance.getCellValue(this.hfInstance.simpleCellAddressFromString(cell, sheet)); // works after reset
      console.log(cell, result);
    }
    else {
      console.log(cell, result);
    }
  }



}
