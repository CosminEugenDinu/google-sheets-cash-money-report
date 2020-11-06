


class Range{
  constructor(row_i, col_i, numOfRows, numOfCols){
    this.row_i = row_i;
    this.col_i = col_i;
    this.numOfRows = numOfRows;
    this.numOfCols = numOfCols;
    this.values = [];
  }

  getValues(){
    const values = [];
    for (let i=this.row_i; i<this.numOfRows; i++){
      const row = [];
      for (let j=this.col_i; j<this.numOfCols; j++){
        row.push(this.values[i][j]);
      }
      values.push(row);
    }
    return values;
  }

  setValues(newValues){
    if (newValues.length > this.values.length){
      //throw new Error(`newValues.length ${newValues.length} > values.length${this.values.length}`);
      let numOfNew = this.values.length - newValues.length;
      while (numOfNew-- > 0){
        this.values.push(Array(newValues[0].length));
      }
    }
    let row_i = this.row_i; 
    for (let i=0; i<newValues.length; i++){
      this.values[row_i++] = newValues[i];
    }
  }

  clear(){
    return this;
  }
}

class Sheet{
  constructor(){
    this.values = [];
  }

  getRange(...args){
    // if is only one arg, like 'A1:F'
    if (args.length === 1 && typeof args[0] === 'string'){
      const annot = args[0];
      const colLetters = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q'];
      const reAnnot = /([A-Z]*)(\d*):([A-Z]*)(\d*)/;
      let [_, fromCol, fromRow, toCol, toRow] = annot.match(reAnnot);
      fromCol = colLetters.indexOf(fromCol)+1;
      fromRow = fromRow ? +fromRow : 1;
      toCol = colLetters.indexOf(toCol)+2;
      toRow = toRow ? +toRow+1 : this.values.length+1;
      
      const [fromRow_i, fromCol_i, numOfRows, numOfCols] =
        [fromRow-1, fromCol-1, toRow-fromRow, toCol-fromCol];
      //throw `${fromRow}, ${fromCol}, ${numOfRows}, ${numOfCols}`;

      const newRange = new Range(fromRow_i, fromCol_i, numOfRows, numOfCols);
      newRange.values = this.values;
      return newRange;

    } else if (args.length === 4){
      const [row, col, numOfRows, numOfCols] = args;
      const newRange = new Range(row-1, col-1, numOfRows, numOfCols);
      newRange.values = this.values;
      return newRange;
    }
  }
}

module.exports = Sheet;

