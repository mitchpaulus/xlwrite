//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     ANTLR Version: 4.13.1
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

// Generated from XlWrite.g4 by ANTLR 4.13.1

// Unreachable code detected
#pragma warning disable 0162
// The variable '...' is assigned but its value is never used
#pragma warning disable 0219
// Missing XML comment for publicly visible type or member '...'
#pragma warning disable 1591
// Ambiguous reference in cref attribute
#pragma warning disable 419

using System;
using System.IO;
using System.Text;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Misc;
using DFA = Antlr4.Runtime.Dfa.DFA;

[System.CodeDom.Compiler.GeneratedCode("ANTLR", "4.13.1")]
[System.CLSCompliant(false)]
public partial class XlWriteLexer : Lexer {
	protected static DFA[] decisionToDFA;
	protected static PredictionContextCache sharedContextCache = new PredictionContextCache();
	public const int
		T__0=1, T__1=2, T__2=3, T__3=4, T__4=5, T__5=6, T__6=7, T__7=8, CELL=9, 
		COLON=10, STRING=11, LCURLY=12, RCURLY=13, COMMA=14, INT=15, WS=16;
	public static string[] channelNames = {
		"DEFAULT_TOKEN_CHANNEL", "HIDDEN"
	};

	public static string[] modeNames = {
		"DEFAULT_MODE"
	};

	public static readonly string[] ruleNames = {
		"T__0", "T__1", "T__2", "T__3", "T__4", "T__5", "T__6", "T__7", "CELL", 
		"COLON", "STRING", "ESC", "LCURLY", "RCURLY", "COMMA", "INT", "WS"
	};


	public XlWriteLexer(ICharStream input)
	: this(input, Console.Out, Console.Error) { }

	public XlWriteLexer(ICharStream input, TextWriter output, TextWriter errorOutput)
	: base(input, output, errorOutput)
	{
		Interpreter = new LexerATNSimulator(this, _ATN, decisionToDFA, sharedContextCache);
	}

	private static readonly string[] _LiteralNames = {
		null, "'bold'", "'fill'", "'width'", "'border'", "'red'", "'blue'", "'black'", 
		"'rgb'", null, "':'", null, "'{'", "'}'", "','"
	};
	private static readonly string[] _SymbolicNames = {
		null, null, null, null, null, null, null, null, null, "CELL", "COLON", 
		"STRING", "LCURLY", "RCURLY", "COMMA", "INT", "WS"
	};
	public static readonly IVocabulary DefaultVocabulary = new Vocabulary(_LiteralNames, _SymbolicNames);

	[NotNull]
	public override IVocabulary Vocabulary
	{
		get
		{
			return DefaultVocabulary;
		}
	}

	public override string GrammarFileName { get { return "XlWrite.g4"; } }

	public override string[] RuleNames { get { return ruleNames; } }

	public override string[] ChannelNames { get { return channelNames; } }

	public override string[] ModeNames { get { return modeNames; } }

	public override int[] SerializedAtn { get { return _serializedATN; } }

	static XlWriteLexer() {
		decisionToDFA = new DFA[_ATN.NumberOfDecisions];
		for (int i = 0; i < _ATN.NumberOfDecisions; i++) {
			decisionToDFA[i] = new DFA(_ATN.GetDecisionState(i), i);
		}
	}
	private static int[] _serializedATN = {
		4,0,16,123,6,-1,2,0,7,0,2,1,7,1,2,2,7,2,2,3,7,3,2,4,7,4,2,5,7,5,2,6,7,
		6,2,7,7,7,2,8,7,8,2,9,7,9,2,10,7,10,2,11,7,11,2,12,7,12,2,13,7,13,2,14,
		7,14,2,15,7,15,2,16,7,16,1,0,1,0,1,0,1,0,1,0,1,1,1,1,1,1,1,1,1,1,1,2,1,
		2,1,2,1,2,1,2,1,2,1,3,1,3,1,3,1,3,1,3,1,3,1,3,1,4,1,4,1,4,1,4,1,5,1,5,
		1,5,1,5,1,5,1,6,1,6,1,6,1,6,1,6,1,6,1,7,1,7,1,7,1,7,1,8,4,8,79,8,8,11,
		8,12,8,80,1,8,4,8,84,8,8,11,8,12,8,85,1,9,1,9,1,10,1,10,1,10,5,10,93,8,
		10,10,10,12,10,96,9,10,1,10,1,10,1,11,1,11,1,11,1,11,3,11,104,8,11,1,12,
		1,12,1,13,1,13,1,14,1,14,1,15,4,15,113,8,15,11,15,12,15,114,1,16,4,16,
		118,8,16,11,16,12,16,119,1,16,1,16,1,94,0,17,1,1,3,2,5,3,7,4,9,5,11,6,
		13,7,15,8,17,9,19,10,21,11,23,0,25,12,27,13,29,14,31,15,33,16,1,0,3,2,
		0,65,90,97,122,1,0,48,57,3,0,9,10,13,13,32,32,128,0,1,1,0,0,0,0,3,1,0,
		0,0,0,5,1,0,0,0,0,7,1,0,0,0,0,9,1,0,0,0,0,11,1,0,0,0,0,13,1,0,0,0,0,15,
		1,0,0,0,0,17,1,0,0,0,0,19,1,0,0,0,0,21,1,0,0,0,0,25,1,0,0,0,0,27,1,0,0,
		0,0,29,1,0,0,0,0,31,1,0,0,0,0,33,1,0,0,0,1,35,1,0,0,0,3,40,1,0,0,0,5,45,
		1,0,0,0,7,51,1,0,0,0,9,58,1,0,0,0,11,62,1,0,0,0,13,67,1,0,0,0,15,73,1,
		0,0,0,17,78,1,0,0,0,19,87,1,0,0,0,21,89,1,0,0,0,23,103,1,0,0,0,25,105,
		1,0,0,0,27,107,1,0,0,0,29,109,1,0,0,0,31,112,1,0,0,0,33,117,1,0,0,0,35,
		36,5,98,0,0,36,37,5,111,0,0,37,38,5,108,0,0,38,39,5,100,0,0,39,2,1,0,0,
		0,40,41,5,102,0,0,41,42,5,105,0,0,42,43,5,108,0,0,43,44,5,108,0,0,44,4,
		1,0,0,0,45,46,5,119,0,0,46,47,5,105,0,0,47,48,5,100,0,0,48,49,5,116,0,
		0,49,50,5,104,0,0,50,6,1,0,0,0,51,52,5,98,0,0,52,53,5,111,0,0,53,54,5,
		114,0,0,54,55,5,100,0,0,55,56,5,101,0,0,56,57,5,114,0,0,57,8,1,0,0,0,58,
		59,5,114,0,0,59,60,5,101,0,0,60,61,5,100,0,0,61,10,1,0,0,0,62,63,5,98,
		0,0,63,64,5,108,0,0,64,65,5,117,0,0,65,66,5,101,0,0,66,12,1,0,0,0,67,68,
		5,98,0,0,68,69,5,108,0,0,69,70,5,97,0,0,70,71,5,99,0,0,71,72,5,107,0,0,
		72,14,1,0,0,0,73,74,5,114,0,0,74,75,5,103,0,0,75,76,5,98,0,0,76,16,1,0,
		0,0,77,79,7,0,0,0,78,77,1,0,0,0,79,80,1,0,0,0,80,78,1,0,0,0,80,81,1,0,
		0,0,81,83,1,0,0,0,82,84,7,1,0,0,83,82,1,0,0,0,84,85,1,0,0,0,85,83,1,0,
		0,0,85,86,1,0,0,0,86,18,1,0,0,0,87,88,5,58,0,0,88,20,1,0,0,0,89,94,5,34,
		0,0,90,93,3,23,11,0,91,93,9,0,0,0,92,90,1,0,0,0,92,91,1,0,0,0,93,96,1,
		0,0,0,94,95,1,0,0,0,94,92,1,0,0,0,95,97,1,0,0,0,96,94,1,0,0,0,97,98,5,
		34,0,0,98,22,1,0,0,0,99,100,5,92,0,0,100,104,5,34,0,0,101,102,5,92,0,0,
		102,104,5,92,0,0,103,99,1,0,0,0,103,101,1,0,0,0,104,24,1,0,0,0,105,106,
		5,123,0,0,106,26,1,0,0,0,107,108,5,125,0,0,108,28,1,0,0,0,109,110,5,44,
		0,0,110,30,1,0,0,0,111,113,7,1,0,0,112,111,1,0,0,0,113,114,1,0,0,0,114,
		112,1,0,0,0,114,115,1,0,0,0,115,32,1,0,0,0,116,118,7,2,0,0,117,116,1,0,
		0,0,118,119,1,0,0,0,119,117,1,0,0,0,119,120,1,0,0,0,120,121,1,0,0,0,121,
		122,6,16,0,0,122,34,1,0,0,0,8,0,80,85,92,94,103,114,119,1,6,0,0
	};

	public static readonly ATN _ATN =
		new ATNDeserializer().Deserialize(_serializedATN);


}
