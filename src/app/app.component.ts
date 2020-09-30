import { Component } from '@angular/core';
import { CellType, HyperFormula, SimpleCellAddress } from 'hyperformula';
import { AbsoluteCellRange } from 'hyperformula/typings/AbsoluteCellRange';
import sockeldata from '../assets/_deltatetraneu.json';
import traverse from 'traverse';

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
  hfInstance: HyperFormula;
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
      //   console.log('sheet ' + name, cells);
      for (var x = 0; x < cells.length; x++) {
        for (var y = 0; y < cells[x].length; y++) {
          if (cells[x][y] === "")
            cells[x][y] = null;
        }
      }
      hfInstance.addSheet(name);
      hfInstance.setSheetContent(name, cells);
    }



    // new develop test
    console.log("TEST NEW DEV BRANCH");
    console.log(hfInstance.getCellValue({ sheet: 0, col: 5, row: 2 })); // works, static value
    console.log(hfInstance.getCellValue({ sheet: 0, col: 5, row: 3 })); // #ref
    console.log(hfInstance.getCellValue({ sheet: 0, col: 5, row: 4 })); // #ref
    console.log(hfInstance.getCellValue({ sheet: 0, col: 5, row: 5 })); // #ref
    // with workaround it works
    this.printCellData("F3", 0);
    this.printCellData("F4", 0);
    this.printCellData("F5", 0);
    this.printCellData("F6", 0);






    // for (var i = 0; i < sheetnames.length; i++) {
    //   var name = sheetnames[i];

    //   var array = (sockeldata as any).sheets[name];
    //   const cells = array[0].map((_, colIndex) => array.map(row => row[colIndex]));
    // }

    var idxRahmeneckeNord = hfInstance.getSheetId("RahmeneckeNord");
    console.log('cell', hfInstance.getCellValue({ sheet: idxRahmeneckeNord, col: 0, row: 0 })); // =IF($B$361,0,C526)


    console.log('cell', hfInstance.getCellValue({ sheet: idxRahmeneckeNord, col: 1, row: 310 })); // =RahmeneckeNord!C530
    console.log('cell', hfInstance.getCellValue({ sheet: idxRahmeneckeNord, col: 3, row: 313 })); // =IF(AND(B314=1,C318="Passbolzen"),1,0
    var positionen = hfInstance.getSheetId("Positionen");
    var idxDT = hfInstance.getSheetId("DeltaTetra");
    var ST = hfInstance.getSheetId("Statik Tetra");
    var SD = hfInstance.getSheetId("Statik Delta");
    var auswahl = hfInstance.getSheetId("Auswahl");
    var materialmapping = hfInstance.getSheetId("MaterialMapping");



    console.log('cell b2', hfInstance.getCellValue({ sheet: SD, col: 1, row: 1 })); // =IF($B$361,0,C526)
    console.log('cell b2', this.hfInstance.getCellValue(this.hfInstance.simpleCellAddressFromString("B2", SD)));
    var cellcontent = this.hfInstance.getCellValue(this.hfInstance.simpleCellAddressFromString("B2", SD)); // H169 ='Statik Tetra'!C44
    console.log("B2", cellcontent);
    this.printCellData("Z1375", idxDT);
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
    // for (var i = 7; i < 160; i++) {
    //   hfInstance.setCellContents(hfInstance.simpleCellAddressFromString("J" + i, idxDT), null)
    //   this.printCellData("J" + i, idxDT);
    // }
    // //   
    //  hfInstance.rebuildAndRecalculate();
    this.printCellData("K169", idxDT);

    this.printCellData("K40", idxDT);

    this.printCellData("F24", SD);
    this.printCellData("A2", positionen); this.printCellData("A3", ST);


    this.printCellData("I172", idxDT);

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

    this.printCellData("K1223", idxDT);
    this.printCellData("L1223", idxDT);
    this.printCellData("M1223", idxDT);
    this.printCellData("N1223", idxDT);
    this.printCellData("O1223", idxDT);


    cellcontent = '=SUM(L168:L171)';

    // this.printCellData("L168", idxDT);
    // this.printCellData("L169", idxDT);
    // this.printCellData("L170", idxDT);
    // this.printCellData("L171", idxDT);
    this.printCellData("L172", idxDT);

    this.printCellData("I172", idxDT);

    // this.printCellData("P1223", idxDT);

    // this.printCellData("Q1223", idxDT);
    // this.printCellData("R1223", idxDT);
    // this.printCellData("S1223", idxDT);
    // this.printCellData("T1223", idxDT);

    // cellcontent = '=C1223&";"&D1223&";"&E1223&";"&F1223&";"&G1223&H1223&":"&I1223&":"&J1223&"/"&G1224&H1224&":"&I1224&":"&J1224&"/"&G1225&H1225&":"&I1225&":"&J1225&"/"&G1226&H1226&":"&I1226&":"&J1226&"/"&G1227&H1227&":"&I1227&":"&J1227&"/"&G1228&H1228&":"&I1228&":"&J1228&"/"&G1229&H1229&":"&I1229&":"&J1229&";"&K1223&":"&L1223&":"&M1223&";"&N1223&";"&O1223&";"&P1223&";"&Q1223&";"&R1223&";"&S1223&";"&T1223&";"&U1223&";"&V1223&";"&W1223&";"&X1223&";"&Y1223';
    // this.hfInstance.setCellContents(this.hfInstance.simpleCellAddressFromString("Z1223", idxDT), cellcontent);
    // this.printCellData("Z1223", idxDT);

    // this.printCellData("Z1199", idxDT);
    // this.printCellData("Z1200", idxDT);
    // this.printCellData("Z1201", idxDT);
    // this.printCellData("Z1202", idxDT);
    // this.printCellData("Z1203", idxDT);
    // this.printCellData("Z1204", idxDT);
    // this.printCellData("Z1205", idxDT);
    // this.printCellData("Z1206", idxDT);
    // this.printCellData("Z1207", idxDT);
    // this.printCellData("Z1208", idxDT);
    // this.printCellData("Z1209", idxDT);

    var json = {
      parameters: {
        Template: "WEB-VORLAGE-02",
        HalleBreiteH1: "2000",
        HalleLängeH1: "5000",
        HallleTraufhoheH1: "7"
      },
      objects: [
        {
          id: "2825c09b-5aa8-425c-965b-5b61c37099da-Achsen-232-25",
          typ: "TEXTBLOCK",
          konfigurationId: "2825c09b-5aa8-425c-965b-5b61c37099da",
          name: "ACHSENBESCHRIFTUNG",
          baugruppe: "Text",
          bauuntergruppe: "",
          text: "Nord",
          textHeight: 150.0,
          info: "",
          material: "",
          points: [
            {
              x: 2500.0,
              y: 2650.0,
              z: 0.0
            }
          ],
          extrusionsvektor: {
            x: 0.0,
            y: 0.0,
            z: 0.0
          },
          booleanOperation: "",
          visible: true,
          verknuepfteElemente: [],
          verknuepfteKonfigurationen: "2825c09b-5aa8-425c-965b-5b61c37099da"
        }
      ]
    };
    // for (var i = 1198; i < 1412; i++) {
    //   this.printCellData("Z" + i, idxDT);
    // }

    // this.printCellData("K109", idxDT);

    // this.printCellData("L1231", idxDT);

    // // cellcontent = '=C1231&";"&D1231&";"&E1231&";"&F1231&";"&G1231&H1231&":"&I1231&":"&J1231&";"&K1231&":"&L1231&":"';//&M1231&";"&N1231&";"&O1231&";"&P1231&";"&Q1231&";"&R1231&";"&S1231&";"&T1231&";"&U1231&";"&V1231&";"&W1231&";"&X1231&";"&Y1231';
    // // this.hfInstance.setCellContents(this.hfInstance.simpleCellAddressFromString("Z1231", idxDT), cellcontent);
    // // this.printCellData("Z1231", idxDT);
    // // for (var i = 2; i < 2000; i++) {
    // //   this.printCellData("A" + i, positionen);
    // //   this.printCellData("B" + i, positionen);
    // //   this.printCellData("C" + i, positionen);
    // //   this.printCellData("D" + i, positionen);
    // //   this.printCellData("E" + i, positionen);
    // //   this.printCellData("F" + i, positionen);
    // //   this.printCellData("G" + i, positionen);
    // // }
    // this.printCellData("K93", idxDT);

    // this.printCellData("P91", idxDT);
    // //   cellcontent = '=VLOOKUP(K93,Positionen!A2:G2000,7,FALSE)';
    // //    this.hfInstance.setCellContents(this.hfInstance.simpleCellAddressFromString("P91", idxDT), cellcontent);
    // this.printCellData("P91", idxDT);


    // cellcontent = '=\"[R\"&11/2&\"]\"';

    // //this.hfInstance.setCellContents(this.hfInstance.simpleCellAddressFromString("G1375", idxDT), cellcontent);
    // this.printCellData("G1375", idxDT);


    // cellcontent = '=1375&";"&D1375&";"&E1375&";"&F1375&";"&G1375&H1375&":"&I1375&":"&J1375&";"&K1375&":"&L1375&":"&M1375&";"&N1375&";"&O1375&";"&P1375&";"&Q1375&";"&R1375&";"&S1375&";"&T1375&";"&U1375&";"&V1375&";"&W1375&";"&X1375&";"&Y1375';
    // this.hfInstance.setCellContents(this.hfInstance.simpleCellAddressFromString("Z1375", idxDT), cellcontent);
    // this.printCellData("Z1375", idxDT);

    // this.printCellData("I172", idxDT);
    // this.printCellData("Z1218", idxDT);
    // this.printCellData("Z1231", idxDT);
    // this.printCellData("Z1251", idxDT);
    // this.printCellData("Z1280", idxDT);
    // this.printCellData("Z1286", idxDT);
    // this.printCellData("Z1412", idxDT);
    // this.printCellData("Z1413", idxDT);
    this.createJsonMenu();
  }

  getStartCell3D() {
    var sheetv3d = this.hfInstance.getSheetId("View-3d");
    var startrow = this.hfInstance.getCellValue({ sheet: sheetv3d, col: 2, row: 2 });
    var startcol = this.hfInstance.getCellValue({ sheet: sheetv3d, col: 3, row: 2 });
    var sheetname = this.hfInstance.getCellValue({ sheet: sheetv3d, col: 1, row: 2 });

    return { sheet: sheetname, col: startcol, row: startrow };
  }


  createJsonMenu() {
    var sheetname = this.getStartCell3D().sheet.toString();
    var sheetID = this.hfInstance.getSheetId(sheetname);

    // root node
    var controlname = this.hfInstance.getCellValue({ sheet: sheetID, col: 0, row: 3 });
    var controlvisible = this.hfInstance.getCellValue({ sheet: sheetID, col: 2, row: 3 });
    var controlparent = this.hfInstance.getCellValue({ sheet: sheetID, col: 3, row: 3 });
    var controltyp = this.hfInstance.getCellValue({ sheet: sheetID, col: 5, row: 3 });
    var controlreadonly = this.hfInstance.getCellValue({ sheet: sheetID, col: 6, row: 3 });
    var controllabel = this.hfInstance.getCellValue({ sheet: sheetID, col: 7, row: 3 });
    var controloutput = this.hfInstance.getCellValue({ sheet: sheetID, col: 9, row: 3 });
    var controlmin = this.hfInstance.getCellValue({ sheet: sheetID, col: 11, row: 3 });
    var controlmax = this.hfInstance.getCellValue({ sheet: sheetID, col: 12, row: 3 });
    var controlunit = this.hfInstance.getCellValue({ sheet: sheetID, col: 13, row: 3 });

    var menu = {
      "parameters": {},
      "menu": {
        "controlName": "MainTab",
        groupboxTyp: "Tab",
        "label": "",
        "infotext": "",
        "typ": "Groupbox",
        "isVisible": true,
        "isReadOnly": false,
        "outputValue": "",
        "outputVarianteValue": "",
        "minValue": "",
        "maxValue": "",
        "nbrDecimal": null,
        "childs": []
      }
    };

    // iterate over rows and attach child nodes
    var row = 3;
    while (this.hfInstance.getCellValue({ sheet: sheetID, col: 0, row: row }) !== "END") {
      var controlname = this.hfInstance.getCellValue({ sheet: sheetID, col: 0, row: row });
      var controlvisible = this.hfInstance.getCellValue({ sheet: sheetID, col: 2, row: row });
      var controlparent = this.hfInstance.getCellValue({ sheet: sheetID, col: 3, row: row });
      var controltyp = this.hfInstance.getCellValue({ sheet: sheetID, col: 5, row: row });
      var controlreadonly = this.hfInstance.getCellValue({ sheet: sheetID, col: 6, row: row });
      var controllabel = this.hfInstance.getCellValue({ sheet: sheetID, col: 7, row: row });
      var controloutput = this.hfInstance.getCellValue({ sheet: sheetID, col: 9, row: row });
      var controlmin = this.hfInstance.getCellValue({ sheet: sheetID, col: 11, row: row });
      var controlmax = this.hfInstance.getCellValue({ sheet: sheetID, col: 12, row: row });
      var controlunit = this.hfInstance.getCellValue({ sheet: sheetID, col: 13, row: row });
      var controlauswahl = this.hfInstance.getCellValue({ sheet: sheetID, col: 14, row: row });

      var newnode = {
        controlName: controlname,
        label: controllabel,
        infotext: "",
        groupboxTyp: controltyp,
        typ: controltyp,
        isVisible: controlvisible,
        isReadOnly: controlreadonly,
        outputValue: controloutput,
        outputVarianteValue: "",
        minValue: controlmin,
        maxValue: controlmax,
        parent: controlparent,
        nbrDecimal: null,
        auswahl: controlauswahl,
        selectOptions: [],
        childs: []
      }

      if (controlauswahl) {

        console.log(row, controlauswahl, newnode);
        var pres = this.hfInstance.getCellPrecedents({ sheet: sheetID, col: 14, row: row })[0] as any;
        //  console.log(controlauswahl, pres);
        //      console.log(this.hfInstance.getCellValue(pres as any));

        pres.row += 2;
        for (var i = 0; i < 10; i++) {
          pres.row += 1;
          //          console.log(pres)
          var c = this.hfInstance.simpleCellAddressToString(pres, pres.sheet)
          // this.printCellData(c, pres.sheet)
          var val = this.hfInstance.getCellValue(pres as any);
          console.log('value', val);
          // if (!val)
          //   break;
          // else {
          //   newnode.selectOptions.push(val);
          //   }
        }
      }
      row += 1;

      var nm = menu;
      traverse(menu).map((value) => {
        try {

          if (typeof (value) === "object" && value && value.controlName) {

            if (value.controlName === newnode.parent) {
              console.log(value.path)
              //              value.path
              value.childs.push(newnode)
              //              this.update(value)
              // value.childs.push(newnode);
              // console.log(menu)
              return;
            }
            if (value.childs) {

              // if (value.childs[Object.keys(value.childs)[0]]) {
              //   //   console.log(value.childs[Object.keys(value.childs)[i]].groupboxTyp)
              //   //  console.log(Object.keys(value.childs).length)
              //   for (var i = 0; i < Object.keys(value.childs).length; i++) {
              //     if (value.childs[Object.keys(value.childs)[i]].groupboxTyp == "Accordium") {
              //       value.isAccordion = true;
              //       break;
              //     }
              //   }
              // }
            }


          }
        } catch (err) {
          console.log(err);
        }
      });


    }


    console.log(menu);





  }


  printCellData(cell: string, sheet) {
    var cellcontent = this.hfInstance.getCellFormula(this.hfInstance.simpleCellAddressFromString(cell, sheet)); // H169 ='Statik Tetra'!C44
    console.log(cell, cellcontent);
    var result = this.hfInstance.getCellValue(this.hfInstance.simpleCellAddressFromString(cell, sheet));
    if ((result as any)?.value === "#REF!" || (result as any)?.value === "#VALUE!") {
      var t = null;
      var precedents = (this.hfInstance.getCellPrecedents(this.hfInstance.simpleCellAddressFromString(cell, sheet)));
      if (precedents?.length > 0) {
        try {
          t = (this.hfInstance.getCellPrecedents(this.hfInstance.simpleCellAddressFromString(cell, sheet)) as unknown as AbsoluteCellRange)[0].arrayOfAddressesInRange();
          if (t == null) {
            var oneref = (this.hfInstance.getCellPrecedents(this.hfInstance.simpleCellAddressFromString(cell, sheet)) as unknown as AbsoluteCellRange)[0];
            t.push(oneref);
          }



          for (var i = 0; i < t.length; i++) {

            try {
              var sname = this.hfInstance.getSheetName(t[i][0].sheet);
              var c = this.hfInstance.simpleCellAddressToString(t[i][0], t[i][0].sheet);
              this.printCellData(c, t[i][0].sheet);

              // TODO: VLOOKUP 
              // =VLOOKUP(1,$G$168:$K$171,3,FALSE)
              // if (cellcontent.startsWith('=VLOOKUP')) {
              //   var offset = parseInt(cellcontent.split(',')[2]);
              //   for (var j = 0; j < t[i].length; j++) {
              //     var col = t[i][j].col;
              //     var row = t[i][j].row;
              //     var sca = {

              //       col: col + offset,
              //       row: row,
              //       sheet: t[i][j].sheet

              //     };

              //     var c = this.hfInstance.simpleCellAddressToString(sca, sca.sheet);
              //     this.printCellData(c, sca.sheet);
              //   }

              // }

            } catch (err) { console.log(err) }
          }
        }
        catch (e) { }

      }

      this.hfInstance.setCellContents(this.hfInstance.simpleCellAddressFromString(cell, sheet), cellcontent);
      result = this.hfInstance.getCellValue(this.hfInstance.simpleCellAddressFromString(cell, sheet)); // works after reset
      console.log(cell, result);
    }
    else {
      console.log(cell, result);
    }
  }



  create3DJson() {
    const options = {
      headers: {},
      responseType: "text"
    };
    // const txt = await this.http
    //   .get("assets/Fachwerkbinder.txt", { responseType: "text" })
    //   .toPromise()
    //   .then();
    var txt = "";
    const lines = txt.split("\r\n");
    let c = 0;
    var rows = [];

    var d1 = new Date().getTime();

    var meshes = [];
    lines.forEach(l => {
      const s = l.split(";");
      const coords = s[4].split("/");
      const cs = [];
      coords.forEach(c => {
        const co = c.split(":");
        cs.push({
          x: parseFloat(co[0]),
          y: parseFloat(co[1]),
          z: parseFloat(co[2])
        });
      });

      const o = {
        nr: s[0],
        name: s[1],
        typ: s[3],
        coords: cs,
        parent: s[9],
        extrusion: {
          // extrusionsvektor
          x: parseFloat(s[5].split(":")[0]),
          y: parseFloat(s[5].split(":")[1]),
          z: parseFloat(s[5].split(":")[2])
        },
        material: s[11]
      };
      rows.push(o);
      console.log(o);

    });
  }


}
