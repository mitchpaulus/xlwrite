file : (selection actions)* ;

CELL : [a-zA-Z]+[0-9]+ ;

COLON : ':' ;

range : CELL COLON CELL ;

selection : CELL | range ;

LCURLY : '{' ;
RCURLY : '}' ;

actions : LCURLY (action (COMMA action)*)? RCURLY ;

action
    : fillAction
    | widthAction
    ;

fillAction : 'fill' COLOR ;
widthAction : 'width' INT ;

INT : [0-9]+ ;
COLOR : 'red' | 'blue'

WS : [ \t\r\n]+ -> skip ;
