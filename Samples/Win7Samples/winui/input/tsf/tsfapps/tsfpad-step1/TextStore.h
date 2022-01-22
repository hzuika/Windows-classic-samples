#pragma once


class CTextEditor;

// 編集中の文字列のインデックスベースでテキストを編集するインターフェース
// アプリケーションはテキストサービス(IME or 変換エンジン)に対して，テキストストアを通じてドキュメントの参照と編集を行わせる
// アプリケーションによって実装され、TSFマネージャーがTSFのテキストストリームまたはテキストストアを操作するために使用します。
// アプリケーション文字位置（ACP）形式でテキストストアを公開します。
class CTextStore : public ITextStoreACP
{
public:
     CTextStore(CTextEditor *pEditor) 
     {
         _pEditor = pEditor;
         _cRef = 1;
     }
     ~CTextStore() {}


    //
    // IUnknown methods
    //
    // インターフェースを取得する
    STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj);
    STDMETHODIMP_(ULONG) AddRef(void);
    STDMETHODIMP_(ULONG) Release(void);

    //
    // ITextStoreACP
    //
    // シンクを登録する(アドバイズする)
    // ITextStoreACPSinkインターフェイスから新しいアドバイスシンクをインストールするか、既存のアドバイスシンクを変更します。
    STDMETHODIMP    AdviseSink(REFIID riid, IUnknown *punk, DWORD dwMask);
    // シンクを登録解除する(アンアドバイズ)
    // TSFマネージャーからの通知が不要になったことを示すためにアプリケーションによって呼び出されます。
    // TSFマネージャーは、シンクインターフェイスを解放し、通知を停止します。
    STDMETHODIMP    UnadviseSink(IUnknown *punk);
    // テキストサービスがドキュメントのロックを要求した場合には，このメソッドが呼ばれる
    // ドキュメントを変更するためにドキュメントロックを提供するためにTSFマネージャーによって呼び出されます。
    // ITextStoreACPSink :: OnLockGrantedメソッドを呼び出して、ドキュメントロックを作成します。
    STDMETHODIMP    RequestLock(DWORD dwLockFlags, HRESULT *phrSession);
    // ドキュメントのステータスを取得します。
    // ドキュメントのステータスは、TS_STATUS構造を介して返されます。
    STDMETHODIMP    GetStatus(TS_STATUS *pdcs);
    // 指定された開始文字と終了文字の位置が有効かどうかを判別します。
    STDMETHODIMP    QueryInsert(LONG acpInsertStart, LONG acpInsertEnd, ULONG cch, LONG *pacpResultStart, LONG *pacpResultEnd);
    // 現在選択中の文字列の始点と終点のACPを返す
    // ドキュメント内のテキスト選択の文字位置を返します。
    // このメソッドは、複数のテキスト選択をサポートします。
    // 呼び出し元は、このメソッドを呼び出す前に、ドキュメントに対して読み取り専用ロックを設定する必要があります。
    STDMETHODIMP    GetSelection(ULONG ulIndex, ULONG ulCount, TS_SELECTION_ACP *pSelection, ULONG *pcFetched);
    // IME側からテキストが選択される
    // ドキュメント内のテキストを選択します。
    // このメソッドを呼び出す前に、アプリケーションはドキュメントの読み取り/書き込みロックを持っている必要があります。
    STDMETHODIMP    SetSelection(ULONG ulCount, const TS_SELECTION_ACP *pSelection);
    // 指定された範囲の文字列を返す
    // 指定された文字位置のテキストに関する情報を返します。
    // このメソッドは、表示および非表示のテキストを返し、埋め込みデータがテキストに添付されているかどうかを示します。
    STDMETHODIMP    GetText(LONG acpStart, LONG acpEnd, __out_ecount(cchPlainReq) WCHAR *pchPlain, ULONG cchPlainReq, ULONG *pcchPlainOut, TS_RUNINFO *prgRunInfo,
        ULONG ulRunInfoReq, ULONG *pulRunInfoOut, LONG *pacpNext);
    // 指定された範囲の文字列を指定された文字列で置換する
    // テキスト選択を指定された文字位置に設定します。
    STDMETHODIMP    SetText(DWORD dwFlags, LONG acpStart, LONG acpEnd, __in_ecount(cch) const WCHAR *pchText, ULONG cch, TS_TEXTCHANGE *pChange);
    // 指定されたテキスト文字列に関するフォーマットされたテキストデータを返します。
    // 呼び出し元は、このメソッドを呼び出す前に、ドキュメントの読み取り/書き込みロックを持っている必要があります。
    STDMETHODIMP    GetFormattedText(LONG acpStart, LONG acpEnd, IDataObject **ppDataObject);
    // 埋め込みドキュメントを取得します。
    STDMETHODIMP    GetEmbedded(LONG acpPos, REFGUID rguidService, REFIID riid, IUnknown **ppunk);
    // 指定された文字に埋め込みオブジェクトを挿入します。
    STDMETHODIMP    InsertEmbedded(DWORD dwFlags, LONG acpStart, LONG acpEnd, IDataObject *pDataObject, TS_TEXTCHANGE *pChange);
    // ドキュメントでサポートされている属性を取得します。
    STDMETHODIMP    RequestSupportedAttrs(DWORD dwFlags, ULONG cFilterAttrs, const TS_ATTRID *paFilterAttrs);
    // 指定された文字位置のテキスト属性を取得します。
    STDMETHODIMP    RequestAttrsAtPosition(LONG acpPos, ULONG cFilterAttrs, const TS_ATTRID *paFilterAttrs, DWORD dwFlags);
    // 指定された文字位置で遷移するテキスト属性を取得します。
    STDMETHODIMP    RequestAttrsTransitioningAtPosition(LONG acpPos, ULONG cFilterAttrs, const TS_ATTRID *paFilterAttrs, DWORD dwFlags);
    // 属性値で遷移が発生する文字の位置を決定します。
    // チェックする指定された属性は、アプリケーションによって異なります。
    STDMETHODIMP    FindNextAttrTransition(LONG acpStart, LONG acpHalt, ULONG cFilterAttrs, const TS_ATTRID *paFilterAttrs,
        DWORD dwFlags, LONG *pacpNext, BOOL *pfFound, LONG *plFoundOffset);
    // 属性要求メソッドの呼び出しによって返される属性を取得します。
    STDMETHODIMP    RetrieveRequestedAttrs(ULONG ulCount, TS_ATTRVAL *paAttrVals, ULONG *pcFetched);
    // 編集中の文字列の長さを返す
    // ドキュメント内の文字数を返します。
    STDMETHODIMP    GetEndACP(LONG *pacp);
    // 現在アクティブなビューを指定するTsViewCookieデータ型を返します。
    STDMETHODIMP    GetActiveView(TsViewCookie *pvcView);
    // ディスプレイ上の座標に対応するACP(文字列のインデックス)を返す
    // 画面座標のポイントをアプリケーションの文字位置に変換します。
    STDMETHODIMP    GetACPFromPoint(TsViewCookie vcView, const POINT *pt, DWORD dwFlags, LONG *pacp);
    // 指定されたテキスト範囲がディスプレイ上のどこにあるかを返す
    // 指定された文字位置にあるテキストの境界ボックスを画面座標で返します。
    // 呼び出し元は、このメソッドを呼び出す前に、ドキュメントに対して読み取り専用ロックを設定する必要があります。
    STDMETHODIMP    GetTextExt(TsViewCookie vcView, LONG acpStart, LONG acpEnd, RECT *prc, BOOL *pfClipped);
    // テキストストリームがレンダリングされる表示面の境界ボックスの画面座標を返します。
    STDMETHODIMP    GetScreenExt(TsViewCookie vcView, RECT *prc);
    // 現在のドキュメントに対応するウィンドウへのハンドルを返します。
    STDMETHODIMP    GetWnd(TsViewCookie vcView, HWND *phwnd);
    // 指定されたオブジェクトをドキュメントに挿入できるかどうかを示す値を取得します。
    STDMETHODIMP    QueryInsertEmbedded(const GUID *pguidService, const FORMATETC *pFormatEtc, BOOL *pfInsertable);
    // 挿入ポイントまたは選択範囲にテキストを挿入します。
    // 発信者は、テキストを挿入する前に、ドキュメントの読み取り/書き込みロックを持っている必要があります。
    STDMETHODIMP    InsertTextAtSelection(DWORD dwFlags, __in_ecount(cch) const WCHAR *pchText, ULONG cch, LONG *pacpStart,
        LONG *pacpEnd, TS_TEXTCHANGE *pChange);
    // 挿入ポイントまたは選択範囲にIDataObjectオブジェクトを挿入します。
    // このメソッドを呼び出すクライアントは、IDataObjectオブジェクトをドキュメントに挿入する前に、読み取り/書き込みロックを持っている必要があります。
    STDMETHODIMP    InsertEmbeddedAtSelection(DWORD dwFlags, IDataObject *pDataObject, LONG *pacpStart, 
        LONG *pacpEnd, TS_TEXTCHANGE *pChange);


    void OnSelectionChange()
    {
        if (_pSink)
            _pSink->OnSelectionChange();
    }

    void OnTextChange(LONG acpStart, LONG acpOldEnd, LONG acpNewEnd)
    {
        if (_pSink)
        {
            TS_TEXTCHANGE textChange;
            textChange.acpStart = acpStart;
            textChange.acpOldEnd = acpOldEnd;
            textChange.acpNewEnd = acpNewEnd;
            _pSink->OnTextChange(0, &textChange);
        }
    }

    void OnLayoutChange()
    {
        if (_pSink)
        {
            _pSink->OnLayoutChange(TS_LC_CHANGE, 0x01);
        }
    }

private:
    void PrepareAttributes(ULONG cFilterAttrs, const TS_ATTRID *paFilterAttrs);
   
    CTextEditor *_pEditor;

    TS_ATTRVAL _attrval[8];
    int _nAttrVals;

    // COMの世界でのイベントは，SourceとSinkという概念で表される．
    // イベントソース(Source)はイベントの発信元
    // イベントシンク(Sink)はイベントの受け取り側
    ITextStoreACPSink *_pSink;

    long _cRef;
};

