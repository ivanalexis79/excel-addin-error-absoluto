Office.onReady(() => {
    document.getElementById("btnSeleccionar").addEventListener("click", seleccionarRango);
    document.getElementById("btnCalcular").addEventListener("click", calcularError);
});

async function seleccionarRango() {
    try {
        await Excel.run(async (context) => {
            let range = context.workbook.getSelectedRange();
            range.load("values");
            await context.sync();
            console.log("Datos seleccionados: ", range.values);
        });
    } catch (error) {
        console.error("Error al seleccionar el rango: ", error);
    }
}

async function calcularError() {
    try {
        await Excel.run(async (context) => {
            let range = context.workbook.getSelectedRange();
            range.load("values");
            await context.sync();

            let valores = range.values;
            let resultado = valores.map(row => [Math.abs(row[0] - row[1])]);

            let outputRange = range.getResizedRange(0, 1);
            outputRange.values = resultado;
            await context.sync();

            document.getElementById("resultado").innerText = "Error Absoluto Calculado";
        });
    } catch (error) {
        console.error("Error al calcular: ", error);
    }
}
