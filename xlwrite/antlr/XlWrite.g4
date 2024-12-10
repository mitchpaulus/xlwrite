grammar XlWrite;

file : item* ;

item : selection actions ;
CELL : [a-zA-Z]+[0-9]+ ;

COLON : ':' ;

range : CELL COLON CELL ;

// STRING is the optional worksheet name.
selection : STRING? (CELL | range) ;

STRING : '"' (ESC|.)*? '"' ;
fragment ESC : '\\"'  | '\\\\' ;

LCURLY : '{' ;
RCURLY : '}' ;
COMMA : ',' ;

actions : LCURLY (action (COMMA action)*)? RCURLY ;

action
    : fillAction # fillActionExp
    | widthAction # widthActionExp
    | borderAction # borderActionExp
    | boldAction # boldActionExp
    ;

boldAction : 'bold' ;
fillAction : 'fill' color ;
widthAction : 'width' INT ;
borderAction : 'border' color? ;

INT : [0-9]+ ;

color : rgbColor | knownColor ;

knownColor : 'red' | 'blue' | 'black' ;

rgbColor : 'rgb' INT INT INT ;
WS : [ \t\r\n]+ -> skip ;
