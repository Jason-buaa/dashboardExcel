/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("gruop").onclick = group_shape;
    document.getElementById("move").onclick = group_move;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      let shapes = context.workbook.worksheets.getItem("Sheet1").shapes;
      let rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
      rectangle.left = 100;
      rectangle.top = 100;
      rectangle.height = 150;
      rectangle.width = 150;
      rectangle.name = "Square";
      let textbox = shapes.addTextBox("Hello!");
      textbox.left = 200;
      textbox.top = 100;
      textbox.height = 20;
      textbox.width = 45;
      textbox.name = "Textbox";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function group_shape() {
  try {
    await Excel.run(async (context) => {
      let shapes = context.workbook.worksheets.getItem("Sheet1").shapes;
      let square = shapes.getItem("Square");
      let textbox = shapes.getItem("Textbox");
      let shapeGroup = shapes.addGroup([square, textbox]);
      shapeGroup.name = "Group";
      console.log("Shapes grouped");
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function group_move() {
  try {
    await Excel.run(async (context) => {
      let shapes = context.workbook.worksheets.getItem("Sheet1").shapes;
      let shapeGroup = shapes.getItem("Group");
      shapeGroup.incrementLeft(50);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
