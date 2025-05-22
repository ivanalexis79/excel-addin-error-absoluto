function calcularErrorAbsoluto(range) {
    let valores = range.values;
    let resultado = [];
    
    for (let i = 0; i < valores.length; i++) {
        resultado.push([Math.abs(valores[i][0] - valores[i][1])]);
    }
    
    return resultado;
}

Excel.ScriptLab.addFunction("calcularErrorAbsoluto", calcularErrorAbsoluto);
