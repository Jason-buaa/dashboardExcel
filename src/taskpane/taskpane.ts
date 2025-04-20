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
    document.getElementById("turtle").onclick = drawTurtle;
    document.getElementById("star").onclick = spinStar;
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
export async function drawTurtle() {
  try {
    await Excel.run(async (context) => {
      let shapes = context.workbook.worksheets.getItem("Sheet1").shapes;

      // Clear existing shapes
      shapes.load("items");
      await context.sync();
      shapes.items.forEach((shape) => shape.delete());

      // Turtle shell (ellipse)
      const shell = shapes.addGeometricShape(Excel.GeometricShapeType.ellipse);
      shell.left = 100;
      shell.top = 100;
      shell.width = 200;
      shell.height = 140;
      shell.fill.setSolidColor("#2E8B57"); // green shell
      shell.name = "Shell";

      // Head (ellipse)
      const head = shapes.addGeometricShape(Excel.GeometricShapeType.ellipse);
      head.left = 280;
      head.top = 130;
      head.width = 50;
      head.height = 50;
      head.fill.setSolidColor("#3CB371");
      head.name = "Head";

      // Eyes
      const eyeWhite = shapes.addGeometricShape(Excel.GeometricShapeType.ellipse);
      eyeWhite.left = 305;
      eyeWhite.top = 140;
      eyeWhite.width = 10;
      eyeWhite.height = 10;
      eyeWhite.fill.setSolidColor("#FFFFFF");
      eyeWhite.name = "EyeWhite";

      const pupil = shapes.addGeometricShape(Excel.GeometricShapeType.ellipse);
      pupil.left = 308;
      pupil.top = 143;
      pupil.width = 5;
      pupil.height = 5;
      pupil.fill.setSolidColor("#000000");
      pupil.name = "Pupil";

      // Legs
      const legPositions = [
        { left: 90, top: 90 }, // Front left
        { left: 90, top: 190 }, // Rear left
        { left: 220, top: 90 }, // Front right
        { left: 220, top: 190 }, // Rear right
      ];

      legPositions.forEach((pos, index) => {
        const leg = shapes.addGeometricShape(Excel.GeometricShapeType.ellipse);
        leg.left = pos.left;
        leg.top = pos.top;
        leg.width = 30;
        leg.height = 40;
        leg.fill.setSolidColor("#3CB371");
        leg.name = `Leg${index + 1}`;
      });

      // Tail (isosceles triangle)
      const tail = shapes.addGeometricShape(Excel.GeometricShapeType.triangle);
      tail.left = 85;
      tail.top = 145;
      tail.width = 20;
      tail.height = 20;
      tail.rotation = 270;
      tail.fill.setSolidColor("#3CB371");
      tail.name = "Tail";

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function spinStar() {
  try {
    await Excel.run(async (context) => {
      let shapes = context.workbook.worksheets.getItem("Sheet1").shapes;

      // Add a 5-pointed star
      const star = shapes.addGeometricShape(Excel.GeometricShapeType.star5);
      star.name = "SpinningStar";
      star.left = 100;
      star.top = 100;
      star.width = 100;
      star.height = 100;
      star.fill.setSolidColor("Gold");

      await context.sync();

      // Function to animate rotation
      let angle = 0;
      const rotateStar = async () => {
        angle = (angle + 10) % 360; // Increment angle
        star.rotation = angle;
        await context.sync();
        setTimeout(rotateStar, 100); // Repeat every 100ms
      };

      rotateStar(); // Start animation
    });
  } catch (error) {
    console.error(error);
  }
}
