// ==UserScript==
// @name         Microsoft 365 Chat（Edge Copilot）会話ダウンロード機能追加
// @namespace    https://sho-lab.co.jp/
// @version      0.3
// @description  Edge Copilotの会話履歴を、テキストもしくは画像として保存する機能を追加します。
// @match        https://copilot.cloud.microsoft/chat*
// @match        https://outlook.office.com/hosted/semanticoverview/Users*
// @require      https://cdnjs.cloudflare.com/ajax/libs/html-to-image/1.11.11/html-to-image.min.js
// @grant        GM_xmlhttpRequest
// ==/UserScript==

// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.js

// @match        https://outlook.office.com/hosted/semanticoverview/*
// @grant        GM_addStyle
// @grant        GM_getResourceText
// @require      BULMA_CSS https://cdn.jsdelivr.net/npm/bulma@1.0.2/css/bulma.min.css

(function() {
    'use strict';

    // 外部ライブラリ読み込み
    function loadScript(src) {
        return new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = src;
            s.onload = resolve;
            s.onerror = reject;
            document.head.appendChild(s);
        });
    }

    // html-to-image と html2canvas を読み込む
/*     Promise.all([
        loadScript('https://cdnjs.cloudflare.com/ajax/libs/html-to-image/1.11.11/html-to-image.min.js')
    ]).then(() => {
        // スクリプト読み込み完了後にメイン処理を開始
        main();
    }).catch(err => {
        console.error('スクリプト読み込みエラー:', err);
        alert('スクリプト読み込みエラー:', err);
    }); */

    main();

    function main() {
        function isString(value) {
            return typeof value === "string";
        }
        function nowString() {
            return new Date().toLocaleString()
                            .replace(/\//g,'')
                            .replace(/:/g,'')
                            .replace(' ','_');
        }

        let isIMEOn = false;

        document.addEventListener('compositionstart', () => {
            isIMEOn = true;
            console.info('compositionstart');
        });
        document.addEventListener('compositionend', () => {
            isIMEOn = false;
            console.info('compositionend');
        });

        window.addEventListener('keydown', function(event) {
            if (event.key === 'Enter') {
                console.info('isIMEOn = ' + isIMEOn);
                console.log('Enterキーが押されました');
                if(!isIMEOn) {
                    createShareConversationButton(event);
                    createShareConversationAsImageButton(event);
                }
            }
        });

        async function createShareConversationButton(event) {
            if(document.querySelector('#share-chat')) return;

            let button = document.createElement('button');
            button.textContent = 'save as text';
            button.id = 'share-chat';
            button.style.backgroundColor = '#9b71c8';
            button.style.color = 'white';
            button.style.border = 'none';
            button.style.borderRadius = '5px';
            button.style.position = 'absolute';
            button.style.right = '80px';
            button.style.bottom = '13px';
            event.target.closest('div').appendChild(button);
            let bodyText = event.target.closest('body').textContent;
            console.info(bodyText);
            button.addEventListener('click', async (e) => {
                const chatLogElement = document.querySelector('#llm-web-ui-messageList-scrollable-container');
                await captureAndDownloadAsText(chatLogElement);
            });
        }

        function createShareConversationAsImageButton(event) {
            if(document.querySelector('#share-chat2')) return;

            let button = document.createElement('button');
            button.textContent = 'save as image';
            button.style.right = '178px';
            button.id = 'share-chat2';
            button.style.backgroundColor = '#9b71c8';
            button.style.color = 'white';
            button.style.border = 'none';
            button.style.borderRadius = '5px';
            button.style.position = 'absolute';
            button.style.bottom = '13px';
            event.target.closest('div').appendChild(button);
            let bodyText = event.target.closest('body').textContent;
            button.addEventListener('click', async (e) => {
                const chatLogElement = document.querySelector('#llm-web-ui-messageList-scrollable-container');
                await captureAndDownloadAsPng(chatLogElement);
            });
        }

        window.addEventListener("click", (event) => {
            const text = event.target.closest('div') ? event.target.closest('div').textContent : null;
            if(text) {
                console.log(text);
                // svg(pathタグ)の場合にボタン挿入を試みる
                if(!document.querySelector('#share-chat') && event.target.tagName.toLowerCase() == 'path'){
                    createShareConversationButton(event);
                    createShareConversationAsImageButton(event);
                }
            }
        }, true);


        // 会話履歴をテキストとして取得する
        async function captureAndDownloadAsText(targetElement) {
            let mdText = '';
            document.querySelectorAll('[class^="largeContainer-2"]').forEach( div => {
                div.childNodes.forEach(dom => {
                    if (dom.textContent) {
                        dom.childNodes.forEach(node=>{
                            if(node.nodeName === '#text') {
                                mdText = mdText + node.textContent + '\n';
                            } else {
                                node.childNodes.forEach(node2=>{
                                    if(node2.nodeName === '#text') {
                                        mdText = mdText + node2.textContent + '\n';
                                    } else if(!node2.querySelector || !node2.querySelector('button')){
                                        mdText = mdText + (node2.textContent||'') + '\n';
                                    }
                                });
                            }
                        });
                    }
                    try {
                        if (dom.querySelector && dom.querySelector('[class^="referenceList-"]')) {
                            if(dom.querySelectorAll('[class^="referenceList-"] button').length >= 1){
                                mdText = mdText + '\nリファレンスリスト（参考サイト一覧）：\n';
                            }
                            dom.querySelectorAll('[class^="referenceList-"] button').forEach(button => {
                                let btnText = '';
                                button.childNodes.forEach(node => {
                                    if (node.textContent && node.textContent.trim() != 'さらに表示') {
                                        btnText += node.textContent + ' |';
                                    }
                                });
                                if (btnText) {
                                    mdText = mdText + btnText + '\n';
                                }
                            });
                        }
                    } catch (err) {
                        console.error('Error processing reference list:', err);
                    }
                });
                mdText = mdText + '\n';
            });
            // 不要なテキストを置換
            mdText = mdText.replace(/Pages で編集コピーAI で生成されたコンテンツは誤りを含む可能性があります。BizChat ついてフィードバックする/g, '');

            try {
                // テキストファイルとしてダウンロード
                const blob = new Blob([mdText], { type: 'text/plain;charset=utf-8' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = `chat-conversation-${nowString()}.txt`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
            } catch (error) {
                console.error('Failed to save conversation:', error);
                alert('会話の保存に失敗しました。');
            }
        }

        // スクリーンショット撮影とダウンロード処理
        async function captureAndDownloadAsPng(targetElement) {
            if (!targetElement) {
                console.error('Target element not found');
                return;
            }

            try {
                const originalScrollTop = targetElement.scrollTop;
                const scrollHeight = targetElement.scrollHeight;

                const originalStyle = targetElement.style.cssText;
                targetElement.style.height = 'auto';
                targetElement.style.overflow = 'visible';
                targetElement.style.maxHeight = 'none';

                const dataUrl = await htmlToImage.toPng(targetElement, {
                    backgroundColor: '#ffffff',
                    pixelRatio: 2,
                    height: scrollHeight
                });

                // 元の状態に戻す
                targetElement.style.cssText = originalStyle;
                targetElement.scrollTop = originalScrollTop;

                // ダウンロード
                const link = document.createElement('a');
                link.href = dataUrl;
                link.download = `screenshot-chat-${nowString()}.png`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);

            } catch (error) {
                console.error('Screenshot failed:', error);
            }
        }

    }

})();
