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
using System.Diagnostics;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using DFA = Antlr4.Runtime.Dfa.DFA;

[System.CodeDom.Compiler.GeneratedCode("ANTLR", "4.13.1")]
[System.CLSCompliant(false)]
public partial class XlWriteParser : Parser {
	protected static DFA[] decisionToDFA;
	protected static PredictionContextCache sharedContextCache = new PredictionContextCache();
	public const int
		T__0=1, T__1=2, T__2=3, T__3=4, T__4=5, T__5=6, T__6=7, T__7=8, CELL=9, 
		COLON=10, STRING=11, LCURLY=12, RCURLY=13, COMMA=14, INT=15, WS=16;
	public const int
		RULE_file = 0, RULE_item = 1, RULE_range = 2, RULE_selection = 3, RULE_actions = 4, 
		RULE_action = 5, RULE_boldAction = 6, RULE_fillAction = 7, RULE_widthAction = 8, 
		RULE_borderAction = 9, RULE_color = 10, RULE_knownColor = 11, RULE_rgbColor = 12;
	public static readonly string[] ruleNames = {
		"file", "item", "range", "selection", "actions", "action", "boldAction", 
		"fillAction", "widthAction", "borderAction", "color", "knownColor", "rgbColor"
	};

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

	public override int[] SerializedAtn { get { return _serializedATN; } }

	static XlWriteParser() {
		decisionToDFA = new DFA[_ATN.NumberOfDecisions];
		for (int i = 0; i < _ATN.NumberOfDecisions; i++) {
			decisionToDFA[i] = new DFA(_ATN.GetDecisionState(i), i);
		}
	}

		public XlWriteParser(ITokenStream input) : this(input, Console.Out, Console.Error) { }

		public XlWriteParser(ITokenStream input, TextWriter output, TextWriter errorOutput)
		: base(input, output, errorOutput)
	{
		Interpreter = new ParserATNSimulator(this, _ATN, decisionToDFA, sharedContextCache);
	}

	public partial class FileContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public ItemContext[] item() {
			return GetRuleContexts<ItemContext>();
		}
		[System.Diagnostics.DebuggerNonUserCode] public ItemContext item(int i) {
			return GetRuleContext<ItemContext>(i);
		}
		public FileContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_file; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterFile(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitFile(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitFile(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public FileContext file() {
		FileContext _localctx = new FileContext(Context, State);
		EnterRule(_localctx, 0, RULE_file);
		int _la;
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 29;
			ErrorHandler.Sync(this);
			_la = TokenStream.LA(1);
			while (_la==CELL || _la==STRING) {
				{
				{
				State = 26;
				item();
				}
				}
				State = 31;
				ErrorHandler.Sync(this);
				_la = TokenStream.LA(1);
			}
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class ItemContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public SelectionContext selection() {
			return GetRuleContext<SelectionContext>(0);
		}
		[System.Diagnostics.DebuggerNonUserCode] public ActionsContext actions() {
			return GetRuleContext<ActionsContext>(0);
		}
		public ItemContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_item; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterItem(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitItem(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitItem(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public ItemContext item() {
		ItemContext _localctx = new ItemContext(Context, State);
		EnterRule(_localctx, 2, RULE_item);
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 32;
			selection();
			State = 33;
			actions();
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class RangeContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode[] CELL() { return GetTokens(XlWriteParser.CELL); }
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode CELL(int i) {
			return GetToken(XlWriteParser.CELL, i);
		}
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode COLON() { return GetToken(XlWriteParser.COLON, 0); }
		public RangeContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_range; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterRange(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitRange(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitRange(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public RangeContext range() {
		RangeContext _localctx = new RangeContext(Context, State);
		EnterRule(_localctx, 4, RULE_range);
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 35;
			Match(CELL);
			State = 36;
			Match(COLON);
			State = 37;
			Match(CELL);
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class SelectionContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode CELL() { return GetToken(XlWriteParser.CELL, 0); }
		[System.Diagnostics.DebuggerNonUserCode] public RangeContext range() {
			return GetRuleContext<RangeContext>(0);
		}
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode STRING() { return GetToken(XlWriteParser.STRING, 0); }
		public SelectionContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_selection; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterSelection(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitSelection(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitSelection(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public SelectionContext selection() {
		SelectionContext _localctx = new SelectionContext(Context, State);
		EnterRule(_localctx, 6, RULE_selection);
		int _la;
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 40;
			ErrorHandler.Sync(this);
			_la = TokenStream.LA(1);
			if (_la==STRING) {
				{
				State = 39;
				Match(STRING);
				}
			}

			State = 44;
			ErrorHandler.Sync(this);
			switch ( Interpreter.AdaptivePredict(TokenStream,2,Context) ) {
			case 1:
				{
				State = 42;
				Match(CELL);
				}
				break;
			case 2:
				{
				State = 43;
				range();
				}
				break;
			}
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class ActionsContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode LCURLY() { return GetToken(XlWriteParser.LCURLY, 0); }
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode RCURLY() { return GetToken(XlWriteParser.RCURLY, 0); }
		[System.Diagnostics.DebuggerNonUserCode] public ActionContext[] action() {
			return GetRuleContexts<ActionContext>();
		}
		[System.Diagnostics.DebuggerNonUserCode] public ActionContext action(int i) {
			return GetRuleContext<ActionContext>(i);
		}
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode[] COMMA() { return GetTokens(XlWriteParser.COMMA); }
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode COMMA(int i) {
			return GetToken(XlWriteParser.COMMA, i);
		}
		public ActionsContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_actions; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterActions(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitActions(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitActions(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public ActionsContext actions() {
		ActionsContext _localctx = new ActionsContext(Context, State);
		EnterRule(_localctx, 8, RULE_actions);
		int _la;
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 46;
			Match(LCURLY);
			State = 55;
			ErrorHandler.Sync(this);
			_la = TokenStream.LA(1);
			if ((((_la) & ~0x3f) == 0 && ((1L << _la) & 30L) != 0)) {
				{
				State = 47;
				action();
				State = 52;
				ErrorHandler.Sync(this);
				_la = TokenStream.LA(1);
				while (_la==COMMA) {
					{
					{
					State = 48;
					Match(COMMA);
					State = 49;
					action();
					}
					}
					State = 54;
					ErrorHandler.Sync(this);
					_la = TokenStream.LA(1);
				}
				}
			}

			State = 57;
			Match(RCURLY);
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class ActionContext : ParserRuleContext {
		public ActionContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_action; } }
	 
		public ActionContext() { }
		public virtual void CopyFrom(ActionContext context) {
			base.CopyFrom(context);
		}
	}
	public partial class BorderActionExpContext : ActionContext {
		[System.Diagnostics.DebuggerNonUserCode] public BorderActionContext borderAction() {
			return GetRuleContext<BorderActionContext>(0);
		}
		public BorderActionExpContext(ActionContext context) { CopyFrom(context); }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterBorderActionExp(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitBorderActionExp(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitBorderActionExp(this);
			else return visitor.VisitChildren(this);
		}
	}
	public partial class WidthActionExpContext : ActionContext {
		[System.Diagnostics.DebuggerNonUserCode] public WidthActionContext widthAction() {
			return GetRuleContext<WidthActionContext>(0);
		}
		public WidthActionExpContext(ActionContext context) { CopyFrom(context); }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterWidthActionExp(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitWidthActionExp(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitWidthActionExp(this);
			else return visitor.VisitChildren(this);
		}
	}
	public partial class FillActionExpContext : ActionContext {
		[System.Diagnostics.DebuggerNonUserCode] public FillActionContext fillAction() {
			return GetRuleContext<FillActionContext>(0);
		}
		public FillActionExpContext(ActionContext context) { CopyFrom(context); }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterFillActionExp(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitFillActionExp(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitFillActionExp(this);
			else return visitor.VisitChildren(this);
		}
	}
	public partial class BoldActionExpContext : ActionContext {
		[System.Diagnostics.DebuggerNonUserCode] public BoldActionContext boldAction() {
			return GetRuleContext<BoldActionContext>(0);
		}
		public BoldActionExpContext(ActionContext context) { CopyFrom(context); }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterBoldActionExp(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitBoldActionExp(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitBoldActionExp(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public ActionContext action() {
		ActionContext _localctx = new ActionContext(Context, State);
		EnterRule(_localctx, 10, RULE_action);
		try {
			State = 63;
			ErrorHandler.Sync(this);
			switch (TokenStream.LA(1)) {
			case T__1:
				_localctx = new FillActionExpContext(_localctx);
				EnterOuterAlt(_localctx, 1);
				{
				State = 59;
				fillAction();
				}
				break;
			case T__2:
				_localctx = new WidthActionExpContext(_localctx);
				EnterOuterAlt(_localctx, 2);
				{
				State = 60;
				widthAction();
				}
				break;
			case T__3:
				_localctx = new BorderActionExpContext(_localctx);
				EnterOuterAlt(_localctx, 3);
				{
				State = 61;
				borderAction();
				}
				break;
			case T__0:
				_localctx = new BoldActionExpContext(_localctx);
				EnterOuterAlt(_localctx, 4);
				{
				State = 62;
				boldAction();
				}
				break;
			default:
				throw new NoViableAltException(this);
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class BoldActionContext : ParserRuleContext {
		public BoldActionContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_boldAction; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterBoldAction(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitBoldAction(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitBoldAction(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public BoldActionContext boldAction() {
		BoldActionContext _localctx = new BoldActionContext(Context, State);
		EnterRule(_localctx, 12, RULE_boldAction);
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 65;
			Match(T__0);
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class FillActionContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public ColorContext color() {
			return GetRuleContext<ColorContext>(0);
		}
		public FillActionContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_fillAction; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterFillAction(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitFillAction(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitFillAction(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public FillActionContext fillAction() {
		FillActionContext _localctx = new FillActionContext(Context, State);
		EnterRule(_localctx, 14, RULE_fillAction);
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 67;
			Match(T__1);
			State = 68;
			color();
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class WidthActionContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode INT() { return GetToken(XlWriteParser.INT, 0); }
		public WidthActionContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_widthAction; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterWidthAction(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitWidthAction(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitWidthAction(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public WidthActionContext widthAction() {
		WidthActionContext _localctx = new WidthActionContext(Context, State);
		EnterRule(_localctx, 16, RULE_widthAction);
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 70;
			Match(T__2);
			State = 71;
			Match(INT);
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class BorderActionContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public ColorContext color() {
			return GetRuleContext<ColorContext>(0);
		}
		public BorderActionContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_borderAction; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterBorderAction(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitBorderAction(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitBorderAction(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public BorderActionContext borderAction() {
		BorderActionContext _localctx = new BorderActionContext(Context, State);
		EnterRule(_localctx, 18, RULE_borderAction);
		int _la;
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 73;
			Match(T__3);
			State = 75;
			ErrorHandler.Sync(this);
			_la = TokenStream.LA(1);
			if ((((_la) & ~0x3f) == 0 && ((1L << _la) & 480L) != 0)) {
				{
				State = 74;
				color();
				}
			}

			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class ColorContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public RgbColorContext rgbColor() {
			return GetRuleContext<RgbColorContext>(0);
		}
		[System.Diagnostics.DebuggerNonUserCode] public KnownColorContext knownColor() {
			return GetRuleContext<KnownColorContext>(0);
		}
		public ColorContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_color; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterColor(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitColor(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitColor(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public ColorContext color() {
		ColorContext _localctx = new ColorContext(Context, State);
		EnterRule(_localctx, 20, RULE_color);
		try {
			State = 79;
			ErrorHandler.Sync(this);
			switch (TokenStream.LA(1)) {
			case T__7:
				EnterOuterAlt(_localctx, 1);
				{
				State = 77;
				rgbColor();
				}
				break;
			case T__4:
			case T__5:
			case T__6:
				EnterOuterAlt(_localctx, 2);
				{
				State = 78;
				knownColor();
				}
				break;
			default:
				throw new NoViableAltException(this);
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class KnownColorContext : ParserRuleContext {
		public KnownColorContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_knownColor; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterKnownColor(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitKnownColor(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitKnownColor(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public KnownColorContext knownColor() {
		KnownColorContext _localctx = new KnownColorContext(Context, State);
		EnterRule(_localctx, 22, RULE_knownColor);
		int _la;
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 81;
			_la = TokenStream.LA(1);
			if ( !((((_la) & ~0x3f) == 0 && ((1L << _la) & 224L) != 0)) ) {
			ErrorHandler.RecoverInline(this);
			}
			else {
				ErrorHandler.ReportMatch(this);
			    Consume();
			}
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	public partial class RgbColorContext : ParserRuleContext {
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode[] INT() { return GetTokens(XlWriteParser.INT); }
		[System.Diagnostics.DebuggerNonUserCode] public ITerminalNode INT(int i) {
			return GetToken(XlWriteParser.INT, i);
		}
		public RgbColorContext(ParserRuleContext parent, int invokingState)
			: base(parent, invokingState)
		{
		}
		public override int RuleIndex { get { return RULE_rgbColor; } }
		[System.Diagnostics.DebuggerNonUserCode]
		public override void EnterRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.EnterRgbColor(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override void ExitRule(IParseTreeListener listener) {
			IXlWriteListener typedListener = listener as IXlWriteListener;
			if (typedListener != null) typedListener.ExitRgbColor(this);
		}
		[System.Diagnostics.DebuggerNonUserCode]
		public override TResult Accept<TResult>(IParseTreeVisitor<TResult> visitor) {
			IXlWriteVisitor<TResult> typedVisitor = visitor as IXlWriteVisitor<TResult>;
			if (typedVisitor != null) return typedVisitor.VisitRgbColor(this);
			else return visitor.VisitChildren(this);
		}
	}

	[RuleVersion(0)]
	public RgbColorContext rgbColor() {
		RgbColorContext _localctx = new RgbColorContext(Context, State);
		EnterRule(_localctx, 24, RULE_rgbColor);
		try {
			EnterOuterAlt(_localctx, 1);
			{
			State = 83;
			Match(T__7);
			State = 84;
			Match(INT);
			State = 85;
			Match(INT);
			State = 86;
			Match(INT);
			}
		}
		catch (RecognitionException re) {
			_localctx.exception = re;
			ErrorHandler.ReportError(this, re);
			ErrorHandler.Recover(this, re);
		}
		finally {
			ExitRule();
		}
		return _localctx;
	}

	private static int[] _serializedATN = {
		4,1,16,89,2,0,7,0,2,1,7,1,2,2,7,2,2,3,7,3,2,4,7,4,2,5,7,5,2,6,7,6,2,7,
		7,7,2,8,7,8,2,9,7,9,2,10,7,10,2,11,7,11,2,12,7,12,1,0,5,0,28,8,0,10,0,
		12,0,31,9,0,1,1,1,1,1,1,1,2,1,2,1,2,1,2,1,3,3,3,41,8,3,1,3,1,3,3,3,45,
		8,3,1,4,1,4,1,4,1,4,5,4,51,8,4,10,4,12,4,54,9,4,3,4,56,8,4,1,4,1,4,1,5,
		1,5,1,5,1,5,3,5,64,8,5,1,6,1,6,1,7,1,7,1,7,1,8,1,8,1,8,1,9,1,9,3,9,76,
		8,9,1,10,1,10,3,10,80,8,10,1,11,1,11,1,12,1,12,1,12,1,12,1,12,1,12,0,0,
		13,0,2,4,6,8,10,12,14,16,18,20,22,24,0,1,1,0,5,7,85,0,29,1,0,0,0,2,32,
		1,0,0,0,4,35,1,0,0,0,6,40,1,0,0,0,8,46,1,0,0,0,10,63,1,0,0,0,12,65,1,0,
		0,0,14,67,1,0,0,0,16,70,1,0,0,0,18,73,1,0,0,0,20,79,1,0,0,0,22,81,1,0,
		0,0,24,83,1,0,0,0,26,28,3,2,1,0,27,26,1,0,0,0,28,31,1,0,0,0,29,27,1,0,
		0,0,29,30,1,0,0,0,30,1,1,0,0,0,31,29,1,0,0,0,32,33,3,6,3,0,33,34,3,8,4,
		0,34,3,1,0,0,0,35,36,5,9,0,0,36,37,5,10,0,0,37,38,5,9,0,0,38,5,1,0,0,0,
		39,41,5,11,0,0,40,39,1,0,0,0,40,41,1,0,0,0,41,44,1,0,0,0,42,45,5,9,0,0,
		43,45,3,4,2,0,44,42,1,0,0,0,44,43,1,0,0,0,45,7,1,0,0,0,46,55,5,12,0,0,
		47,52,3,10,5,0,48,49,5,14,0,0,49,51,3,10,5,0,50,48,1,0,0,0,51,54,1,0,0,
		0,52,50,1,0,0,0,52,53,1,0,0,0,53,56,1,0,0,0,54,52,1,0,0,0,55,47,1,0,0,
		0,55,56,1,0,0,0,56,57,1,0,0,0,57,58,5,13,0,0,58,9,1,0,0,0,59,64,3,14,7,
		0,60,64,3,16,8,0,61,64,3,18,9,0,62,64,3,12,6,0,63,59,1,0,0,0,63,60,1,0,
		0,0,63,61,1,0,0,0,63,62,1,0,0,0,64,11,1,0,0,0,65,66,5,1,0,0,66,13,1,0,
		0,0,67,68,5,2,0,0,68,69,3,20,10,0,69,15,1,0,0,0,70,71,5,3,0,0,71,72,5,
		15,0,0,72,17,1,0,0,0,73,75,5,4,0,0,74,76,3,20,10,0,75,74,1,0,0,0,75,76,
		1,0,0,0,76,19,1,0,0,0,77,80,3,24,12,0,78,80,3,22,11,0,79,77,1,0,0,0,79,
		78,1,0,0,0,80,21,1,0,0,0,81,82,7,0,0,0,82,23,1,0,0,0,83,84,5,8,0,0,84,
		85,5,15,0,0,85,86,5,15,0,0,86,87,5,15,0,0,87,25,1,0,0,0,8,29,40,44,52,
		55,63,75,79
	};

	public static readonly ATN _ATN =
		new ATNDeserializer().Deserialize(_serializedATN);


}
