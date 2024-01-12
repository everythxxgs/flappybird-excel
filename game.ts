function main(workbook: ExcelScript.Workbook){
  let game = new Game(workbook);
  game.setup();
}
class Game{
  static workbook: ExcelScript.Workbook;
  static death = false;
  static sheet: ExcelScript.Worksheet;

  static HEIGHT = 29;
  static WIDTH = 52;
  static CELL_SIZE = 30;

  static PLAYER_X = 4;
  static PLAYER_FALLING = 3;
  static PLAYER_JUMP_HEIGHT = 4;

  

  static PIPE_WIDTH = 3;
  static PIPE_SPACE = 6;
  static PIPE_SPEED = 1;


  static targetRange: ExcelScript.Range;
  
  static tick = 0;
  static tickLastPipe = 0;

  static score = 0;

  static pipes:Pipe[]=[];
  static player;
  static playerY;

  

  constructor( workbook: ExcelScript.Workbook) {
    Game.workbook = workbook;
  }
  setup(){
    Game.sheet = Game.workbook.getWorksheet("Game");

    Game.targetRange = Game.sheet.getRangeByIndexes(0, 0, Game.HEIGHT*2, Game.WIDTH+3);

    Game.targetRange.getFormat().setRowHeight(Game.CELL_SIZE);
    Game.targetRange.getFormat().setColumnWidth(Game.CELL_SIZE);
    Game.targetRange.setValues(null);
    Game.targetRange.getFormat().getFill().setColor("WHITE")

    Game.sheet.getRangeByIndexes(0, 0, Game.HEIGHT, Game.WIDTH).getFormat().getFill().setColor("CYAN");
    Game.sheet.getRangeByIndexes(Game.HEIGHT, 0, Game.HEIGHT*3, Game.WIDTH).getFormat().getFill().setColor("GREEN");

    Game.player = new Player(Game.sheet, Game.HEIGHT, Game.workbook);


    this.loop();
  }
  loop(){
    while (!Game.death) {
      Game.tick++;
      Game.playerY = Game.player.y
      Game.sheet.getCell(1,1).setValue(Game.score);
      console.log(Game.pipes.length);
      Game.player.update() 

      
      if (Game.tick - Game.tickLastPipe > 40) {
        Game.pipes.push(new Pipe(getRandomInt(Game.HEIGHT - Game.PIPE_SPACE*2 - 1)+Game.PIPE_SPACE+1));
        //Game.pipes.push(new Pipe(14));
        Game.tickLastPipe = Game.tick;
      }
      
      for (let i = 0; i < Game.pipes.length&&!Game.death;i++) {
        Game.pipes[i].update();
      }
    }
    Game.sheet.getCell(10,15).getFormat().setRowHeight(78);
    Game.sheet.getCell(10, 15).getFormat().getFont().setSize(78);
    Game.sheet.getCell(10, 15).setValue("YOU LOOSE");

  }
  

}

function getRandomInt(max:number) {
  return Math.floor(Math.random() * max);
}


class Pipe{
  x = 0;

  WIDTH:number;
  HEIGHT:number;
  POSITION:number;
  createdTick:number;
  player;
  index;
  sheet: ExcelScript.Worksheet;

  lastPosition;

  constructor(POSITION:number){

    this.sheet = Game.sheet;
    this.WIDTH,this.x = Game.WIDTH;
    this.HEIGHT = Game.HEIGHT;
    this.player = Game.player;
    this.createdTick = Game.tick;
    this.POSITION = POSITION;
    this.index = Game.pipes.length;
    
  }
  update(){
    this.lastPosition = this.x;
    this.x-= Game.PIPE_SPEED;
    if (Game.PLAYER_X >= this.x -1 && Game.PLAYER_X<= this.x + Game.PIPE_WIDTH ){
      
      if (Game.playerY <= this.POSITION - Game.PIPE_SPACE || Game.playerY > this.POSITION + Game.PIPE_SPACE){
        Game.death=true;
      }
    }
    if (Game.PLAYER_X == this.x - 1 ) {
        Game.score ++;
    }
    if (this.lastPosition <= - Game.PIPE_WIDTH){ 
      Game.pipes.shift;
      return 1;}
    console.log(this.index+" "+(this.POSITION) + " " + (this.lastPosition) + " ")

    this.sheet.getRangeByIndexes(0, Math.max(this.lastPosition, 0), this.POSITION - Game.PIPE_SPACE, Game.PIPE_WIDTH + Math.min(this.lastPosition, 0)).getFormat().getFill().setColor("cyan");
    this.sheet.getRangeByIndexes(this.POSITION + Game.PIPE_SPACE, Math.max(this.lastPosition, 0), this.HEIGHT - Game.PIPE_SPACE - this.POSITION, Game.PIPE_WIDTH+ Math.min(this.lastPosition, 0)).getFormat().getFill().setColor("cyan");

    if (this.lastPosition <= - Game.PIPE_WIDTH+1) return 1;

    this.sheet.getRangeByIndexes(0, Math.max(this.x, 0), this.POSITION - Game.PIPE_SPACE, Game.PIPE_WIDTH + Math.min(this.x,0)).getFormat().getFill().setColor("yellow");
    this.sheet.getRangeByIndexes(this.POSITION + Game.PIPE_SPACE, Math.max(this.x, 0), this.HEIGHT - Game.PIPE_SPACE - this.POSITION, Game.PIPE_WIDTH + Math.min(this.x, 0)).getFormat().getFill().setColor("yellow");

    



  }


}
class Player{
  y = 0;
  old_y = 0;
  h= 0;
  y_0 = 0;
  help_y = 0;
  sheet: ExcelScript.Worksheet;
  book: ExcelScript.Workbook;
  last_Cell;
  constructor(sheet: ExcelScript.Worksheet, height,book:ExcelScript.Workbook){
    this.sheet = sheet;
    this.h = height;
    this.y_0 = this.h;
    this.h = Math.floor(this.h/2);
    this.y,this.old_y = this.h;
    this.sheet.getCell(this.y,4).getFormat().getFill().setColor("red");
    this.book = book;
  }


  update(){
    this.old_y = this.y;
    if (this.y >= Game.HEIGHT - 1){

      Game.death= true;
      return 1;
    }

    let act_cell = this.book.getSelectedRange().getAddress();
    if (act_cell !== this.last_Cell && this.y > 4)
    {
      this.last_Cell = act_cell;
      this.y-= Game.PLAYER_JUMP_HEIGHT;
    }
    
    this.help_y ++;
    if(this.help_y>=Game.PLAYER_FALLING){
      this.help_y = 0;
      this.y++;
      
    }
    this.redraw();
  }
  redraw(){
    this.sheet.getCell(this.old_y,4).getFormat().getFill().setColor("Cyan");
    this.sheet.getCell(this.y,Game.PLAYER_X).getFormat().getFill().setColor("red");

  }
}
