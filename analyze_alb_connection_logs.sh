#!/usr/bin/env bash
# ==============================================================================
# Script Name : analyze_alb_connection_logs.sh
# Description : AWS ALB S3 接続ログ (connection_log / *.gz) 分析・Excelレポート出力スクリプト
# Target OS   : Red Hat Enterprise Linux 9.8 (RHEL 9.8) / Linux (Bash 5.x + Python 3.9+)
# Dependencies: bash, python3 (RHEL 9.8 標準搭載の標準ライブラリのみで動作・外部pip不要)
#
# 概要:
#   指定された input ディレクトリ配下に配置された複数の ALB 接続ログ (.gz 形式) を読み込み、
#   各出力項目の内容をヘッダにわかりやすく設定した上で、1行ずつ確認できる「接続ログ一覧」
#   および多角的な集計を行った「分析サマリーレポート」「エラー・失敗ログ一覧」
#   「項目定義・エラーコード解説」の4シート構成の Excel (.xlsx) ファイルを出力します。
#
# 使用方法:
#   ./analyze_alb_connection_logs.sh <inputディレクトリ> [出力Excelファイルパス]
#   ./analyze_alb_connection_logs.sh -i <inputディレクトリ> [-o <出力Excelファイルパス>] [--no-sort]
# ==============================================================================

set -euo pipefail

# RHEL 9.8 環境でのロケール・文字コード設定 (UTF-8 を優先、pipefail 安全)
AVAILABLE_LOCALES="$(locale -a 2>/dev/null || true)"
if printf '%s\n' "${AVAILABLE_LOCALES}" | grep -i '^C\.utf-\?8$' >/dev/null 2>&1; then
    export LC_ALL=C.UTF-8
    export LANG=C.UTF-8
elif printf '%s\n' "${AVAILABLE_LOCALES}" | grep -i '^en_US\.utf-\?8$' >/dev/null 2>&1; then
    export LC_ALL=en_US.UTF-8
    export LANG=en_US.UTF-8
fi
export PYTHONIOENCODING=utf-8

SCRIPT_NAME="$(basename "$0")"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

usage() {
    cat <<EOF
================================================================================
AWS ALB 接続ログ (connection_log) 分析・Excelレポート出力ツール (RHEL 9.8 対応)
================================================================================

【使用方法】
  ${SCRIPT_NAME} <inputディレクトリ> [出力Excelファイルパス]
  ${SCRIPT_NAME} -i <inputディレクトリ> [-o <出力Excelファイルパス>] [オプション]

【引数・オプション】
  -i, --input <DIR>       ALB接続ログ (.gz) が配置された入力ディレクトリを指定します。
                          (第1引数として直接パスを指定することも可能です。省略時は ./input)
  -o, --output <PATH>     出力する Excel (.xlsx) ファイルパス、または出力先ディレクトリを指定します。
                          (省略時は ./alb_connection_log_report.xlsx に出力します)
  --no-sort               ログをタイムスタンプ順にソートせず、ファイル読込順のまま出力します。
                          (デフォルトは全ファイルのログを接続日時(UTC)の昇順でソートします)
  -h, --help              このヘルプメッセージを表示します。

【実行例】
  1) input ディレクトリを指定して実行 (カレントに alb_connection_log_report.xlsx を出力):
     ./${SCRIPT_NAME} ./input

  2) 入力ディレクトリと出力Excelファイル名を指定して実行:
     ./${SCRIPT_NAME} /var/log/alb/input /var/log/alb/alb_conn_report.xlsx

  3) オプション形式で指定して実行:
     ./${SCRIPT_NAME} -i ./input -o ./output/report.xlsx
================================================================================
EOF
}

log_info() {
    printf '[%s] [INFO]  %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*"
}

log_warn() {
    printf '[%s] [WARN]  %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >&2
}

log_error() {
    printf '[%s] [ERROR] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >&2
}

# ------------------------------------------------------------------------------
# 引数解析
# ------------------------------------------------------------------------------
INPUT_DIR=""
OUTPUT_PATH=""
SORT_BY_TIME="1"
POSITIONAL_ARGS=()

while [[ $# -gt 0 ]]; do
    case "$1" in
        -i|--input)
            if [[ $# -lt 2 ]]; then
                log_error "オプション $1 にはディレクトリパスを指定してください。"
                usage
                exit 1
            fi
            INPUT_DIR="$2"
            shift 2
            ;;
        -o|--output)
            if [[ $# -lt 2 ]]; then
                log_error "オプション $1 には出力ファイルパスを指定してください。"
                usage
                exit 1
            fi
            OUTPUT_PATH="$2"
            shift 2
            ;;
        --no-sort)
            SORT_BY_TIME="0"
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        --)
            shift
            while [[ $# -gt 0 ]]; do
                POSITIONAL_ARGS+=("$1")
                shift
            done
            ;;
        -*)
            log_error "不明なオプションです: $1"
            usage
            exit 1
            ;;
        *)
            POSITIONAL_ARGS+=("$1")
            shift
            ;;
    esac
done

if [[ -z "${INPUT_DIR}" ]]; then
    if [[ ${#POSITIONAL_ARGS[@]} -ge 1 ]]; then
        INPUT_DIR="${POSITIONAL_ARGS[0]}"
    elif [[ -d "./input" ]]; then
        INPUT_DIR="./input"
    elif [[ -d "${SCRIPT_DIR}/input" ]]; then
        INPUT_DIR="${SCRIPT_DIR}/input"
    else
        log_error "入力ディレクトリが指定されていません。"
        usage
        exit 1
    fi
fi

if [[ -z "${OUTPUT_PATH}" ]]; then
    if [[ ${#POSITIONAL_ARGS[@]} -ge 2 ]]; then
        OUTPUT_PATH="${POSITIONAL_ARGS[1]}"
    else
        OUTPUT_PATH="./alb_connection_log_report.xlsx"
    fi
fi

# ------------------------------------------------------------------------------
# 実行環境・入力ディレクトリの検証
# ------------------------------------------------------------------------------
if ! command -v python3 >/dev/null 2>&1; then
    log_error "python3 コマンドが見つかりません。RHEL 9.8 標準の python3 パッケージをインストールしてください。"
    exit 1
fi

if [[ ! -d "${INPUT_DIR}" ]]; then
    log_error "指定された入力ディレクトリが存在しません: ${INPUT_DIR}"
    exit 1
fi

# 出力パスが既存ディレクトリ、または末尾が / の場合はデフォルトファイル名を付与
if [[ -d "${OUTPUT_PATH}" ]] || [[ "${OUTPUT_PATH}" == */ ]]; then
    OUTPUT_PATH="${OUTPUT_PATH%/}/alb_connection_log_report.xlsx"
fi

# 拡張子 .xlsx の補完
if [[ "${OUTPUT_PATH}" != *.xlsx ]]; then
    OUTPUT_PATH="${OUTPUT_PATH}.xlsx"
fi

OUTPUT_DIR="$(dirname "${OUTPUT_PATH}")"
if [[ ! -d "${OUTPUT_DIR}" ]]; then
    mkdir -p "${OUTPUT_DIR}"
fi

# 対象 .gz ファイルの存在確認 (シンボリックリンク先および大文字 .GZ も考慮)
GZ_COUNT="$(find -L "${INPUT_DIR}" -type f \( -name '*.gz' -o -name '*.GZ' \) 2>/dev/null | wc -l | tr -d ' ')"
if [[ "${GZ_COUNT}" -eq 0 ]]; then
    log_error "指定された入力ディレクトリ配下に .gz ファイルが見つかりません: ${INPUT_DIR}"
    exit 1
fi

log_info "ALB 接続ログ分析を開始します。"
log_info "  入力ディレクトリ : ${INPUT_DIR} (検出 .gz ファイル数: ${GZ_COUNT} 件)"
log_info "  出力Excelファイル: ${OUTPUT_PATH}"

# ------------------------------------------------------------------------------
# Python3 による .gz ログ解析・統計分析・OpenXML (.xlsx) レポート生成
# (RHEL 9.8 標準ライブラリのみで完全動作し、外部パッケージ依存なし)
# ------------------------------------------------------------------------------
python3 - "${INPUT_DIR}" "${OUTPUT_PATH}" "${SORT_BY_TIME}" << 'PYEOF'
import sys
import os
import re
import gzip
import json
import math
import zipfile
from pathlib import Path
from datetime import datetime, timezone, timedelta
from collections import Counter, defaultdict
from xml.sax.saxutils import escape as xml_escape

JST = timezone(timedelta(hours=9), name="JST")

# ==============================================================================
# 1. AWS ALB 接続ログ (Connection Log) 仕様定義・エラーコード辞書
# ==============================================================================
AWS_ERROR_DESCRIPTIONS = {
    "ClientCertMaxChainDepthExceeded": (
        "クライアント証明書チェーンの最大深度超過",
        "クライアントが提示した証明書チェーンの階層数がALBの許容上限を超えています。中間CA証明書の構成を確認してください。"
    ),
    "ClientCertMaxSizeExceeded": (
        "クライアント証明書の最大サイズ超過",
        "クライアント証明書のペイロードサイズが上限を超えています。不要な拡張領域や証明書サイズを確認してください。"
    ),
    "ClientCertCrlHit": (
        "クライアント証明書がCRLにより失効済み",
        "認証局(CA)の証明書失効リスト(CRL)に該当シリアル番号が含まれており、失効した証明書として拒否されました。"
    ),
    "ClientCertCrlProcessingError": (
        "CRL(証明書失効リスト)処理エラー",
        "トラストストアに設定されたCRLの処理中にエラーが発生しました。S3上のCRL形式や有効性を確認してください。"
    ),
    "ClientCertUntrusted": (
        "信頼されていないクライアント証明書",
        "ALBのトラストストア内のCA証明書でクライアント証明書を検証できませんでした。発行元CAがトラストストアに登録されているか確認してください。"
    ),
    "ClientCertNotYetValid": (
        "クライアント証明書の有効期間開始前",
        "証明書の有効開始日時(NotBefore)にまだ達していません。証明書の発行日時設定を確認してください。"
    ),
    "ClientCertExpired": (
        "クライアント証明書の有効期限切れ",
        "証明書の有効終了日時(NotAfter)を過ぎています。新しいクライアント証明書へ更新してください。"
    ),
    "ClientCertTypeUnsupported": (
        "サポート外のクライアント証明書タイプ",
        "ALBでサポートされていない証明書形式または鍵アルゴリズムが使用されています。"
    ),
    "ClientCertInvalid": (
        "無効なクライアント証明書",
        "クライアント証明書の構造または署名が不正であり、X.509証明書として検証できません。"
    ),
    "ClientCertPurposeInvalid": (
        "クライアント証明書の用途(拡張鍵用途)不正",
        "証明書のKeyUsage/ExtendedKeyUsageがクライアント認証(clientAuth)用途として有効ではありません。"
    ),
    "ClientCertRejected": (
        "カスタム検証によるクライアント証明書拒否",
        "アプリケーションまたはカスタムサーバー検証ロジックによってクライアント証明書が拒否されました。"
    ),
    "UnmappedConnectionError": (
        "未マッピングのランタイム接続エラー",
        "TLSハンドシェイク中またはTCP切断等により、特定の証明書エラーに分類されないランタイム接続エラーが発生しました。"
    ),
    "ClientCertIncompatible": (
        "リスナーセキュリティポリシー非互換の証明書",
        "クライアント証明書が、選択されたALBリスナーのTLSセキュリティポリシーと互換性がありません。"
    ),
}

# 詳細一覧シートの各出力項目定義 (全20列)
# (col_key, category_header, header_title_with_desc, col_width, field_name_en, data_type, full_description, dash_condition)
DETAIL_COLUMNS = [
    (
        "seq_no",
        "管理項目: No.",
        "No.\n【通し番号】\n1行ごとの連番",
        8,
        "seq_no (独自付与)",
        "Integer",
        "レポート内で1行ずつ確認しやすくするための通し番号（1からの連番）です。",
        "なし（常に連番が出力されます）"
    ),
    (
        "source_file",
        "管理項目: 出力元ファイル",
        "読込元gzファイル名 (source_file)\n【ログ出力元ファイル:行番号】\nS3から取得した元の.gzファイル名とファイル内行番号",
        38,
        "source_file (独自付与)",
        "String",
        "該当の接続ログ行が記録されていた元の .gz ファイル名およびファイル内の行番号です。",
        "なし"
    ),
    (
        "timestamp",
        "AWS項目 #1: timestamp",
        "接続日時 UTC (timestamp)\n【接続確立/失敗日時(UTC)】\nALBが接続を確立または失敗した協定世界時(ISO8601)",
        29,
        "timestamp (位置: 1)",
        "String (ISO 8601)",
        "ロードバランサーがクライアントとの接続を正常に確立できた、または確立に失敗した日時（ISO 8601形式・UTC）です。",
        "なし"
    ),
    (
        "timestamp_jst",
        "分析補助: 日本標準時(JST)",
        "接続日時 JST (timestamp_jst)\n【接続確立/失敗日時(日本時間)】\nUTCから日本標準時(UTC+9)へ変換した日時",
        25,
        "timestamp_jst (独自付与)",
        "String (YYYY-MM-DD HH:MM:SS.sss)",
        "timestamp(UTC)を日本標準時(JST: UTC+9時間)に変換した日時です。国内運用時の時刻照合に使用します。",
        "タイムスタンプが解析不能な場合は -"
    ),
    (
        "client_ip",
        "AWS項目 #2: client_ip",
        "クライアントIP (client_ip)\n【接続元IPアドレス】\nリクエストを送信したクライアントのIPアドレス",
        18,
        "client_ip (位置: 2)",
        "String (IPv4/IPv6)",
        "接続リクエストを送信したクライアント（または直前のプロキシ）のIPアドレスです。",
        "なし"
    ),
    (
        "client_port",
        "AWS項目 #3: client_port",
        "クライアントポート (client_port)\n【接続元ポート番号】\nリクエスト送信元クライアント側のエフェメラルポート番号",
        22,
        "client_port (位置: 3)",
        "Integer",
        "リクエストを送信したクライアント側の送信元ポート番号です。",
        "不明な場合は -"
    ),
    (
        "listener_port",
        "AWS項目 #4: listener_port",
        "リスナーポート (listener_port)\n【ALB待受ポート番号】\nクライアント接続を受信したALBリスナーのポート(443/80等)",
        22,
        "listener_port (位置: 4)",
        "Integer",
        "クライアントの接続リクエストを受信したロードバランサーリスナーのポート番号（例: 443, 80, 8443）です。",
        "不明な場合は -"
    ),
    (
        "tls_protocol",
        "AWS項目 #5: tls_protocol",
        "TLSプロトコル (tls_protocol)\n【SSL/TLSプロトコルバージョン】\nハンドシェイクで使用されたプロトコル(TLSv1.3/TLSv1.2等)",
        22,
        "tls_protocol (位置: 5)",
        "String",
        "[HTTPSリスナー] ハンドシェイク中に使用されたSSL/TLSプロトコル（例: TLSv1.3, TLSv1.2）です。",
        "SSL/TLS以外のリクエスト（HTTP等）やネゴシエーション前の切断時は -"
    ),
    (
        "tls_cipher",
        "AWS項目 #6: tls_cipher",
        "TLS暗号スイート (tls_cipher)\n【SSL/TLS暗号スイート】\nハンドシェイクで選択された暗号スイート(非TLS時は -)",
        34,
        "tls_cipher (位置: 6)",
        "String",
        "[HTTPSリスナー] ハンドシェイク中にネゴシエーションされたSSL/TLS暗号スイート（例: TLS_AES_128_GCM_SHA256, ECDHE-RSA-AES128-GCM-SHA256）です。",
        "SSL/TLS以外のリクエストや暗号スイート確定前の失敗時は -"
    ),
    (
        "tls_handshake_latency",
        "AWS項目 #7: tls_handshake_latency",
        "TLSハンドシェイク時間[秒] (tls_handshake_latency)\n【ハンドシェイク所要時間(秒)】\nTLSハンドシェイク確立に要した合計秒数(ミリ秒精度)",
        28,
        "tls_handshake_latency (位置: 7)",
        "Double (秒)",
        "[HTTPSリスナー] ハンドシェイクの正常な確立に経過した時間の合計（秒単位・ミリ秒精度、例: 0.015, 4.036）です。",
        "①SSL/TLS以外のリクエストの場合、②ハンドシェイクが正常に確立されなかった(Failed)場合は -"
    ),
    (
        "leaf_client_cert_subject",
        "AWS項目 #8: leaf_client_cert_subject",
        "クライアント証明書サブジェクト (leaf_client_cert_subject)\n【mTLSリーフクライアント証明書Subject】\n提示されたクライアント証明書の識別名DN(非該当は -)",
        46,
        "leaf_client_cert_subject (位置: 8)",
        "String",
        "[HTTPSリスナー] mTLS(相互TLS)でクライアントが提示したリーフクライアント証明書のサブジェクト(DN)です。",
        "①非SSL/TLSリクエスト、②リスナーでmTLSが無効、③サーバーが証明書をロード/解析できない場合は -"
    ),
    (
        "leaf_client_cert_validity",
        "AWS項目 #9: leaf_client_cert_validity",
        "クライアント証明書有効期間 (leaf_client_cert_validity)\n【mTLSクライアント証明書の有効期間(UTC)】\n証明書の開始日時(NotBefore)と終了日時(NotAfter)",
        48,
        "leaf_client_cert_validity (位置: 9)",
        "String",
        "[HTTPSリスナー] リーフクライアント証明書の有効期間（ISO 8601形式の NotBefore と NotAfter）です。",
        "①非SSL/TLSリクエスト、②リスナーでmTLSが無効、③サーバーが証明書をロード/解析できない場合は -"
    ),
    (
        "cert_validity_jst",
        "分析補助: 証明書有効期間(JST)",
        "証明書有効期間・JST表示 (cert_validity_jst)\n【mTLS証明書の有効期間(日本時間)】\nNotBefore〜NotAfterを日本時間(JST)に変換した期間表示",
        42,
        "cert_validity_jst (独自付与)",
        "String",
        "leaf_client_cert_validity の NotBefore と NotAfter を日本時間(JST)に変換して見やすく整形した期間です。",
        "leaf_client_cert_validity が - の場合は -"
    ),
    (
        "leaf_client_cert_serial_number",
        "AWS項目 #10: leaf_client_cert_serial_number",
        "クライアント証明書シリアル番号 (leaf_client_cert_serial_number)\n【mTLSクライアント証明書シリアル番号】\nリーフクライアント証明書固有の16進シリアル番号",
        32,
        "leaf_client_cert_serial_number (位置: 10)",
        "String (Hex)",
        "[HTTPSリスナー] リーフクライアント証明書のシリアル番号（16進文字列）です。",
        "①非SSL/TLSリクエスト、②リスナーでmTLSが無効、③サーバーが証明書をロード/解析できない場合は -"
    ),
    (
        "tls_verify_status",
        "AWS項目 #11: tls_verify_status",
        "TLS接続検証ステータス (tls_verify_status)\n【接続確立・証明書検証結果】\n成功時は Success、失敗時は Failed:<エラーコード>",
        30,
        "tls_verify_status (位置: 11)",
        "String",
        "[HTTPSリスナー] 接続リクエストの検証ステータスです。正常確立時は Success、失敗時は Failed:$error_code が出力されます。",
        "非SSL/TLSリクエスト(HTTPリスナー等)の場合は -"
    ),
    (
        "status_analysis",
        "分析補助: ステータス判定・エラー解説",
        "接続判定・エラー内容解説 (status_analysis)\n【接続結果の日本語判定とエラー原因説明】\nSuccess/Failedの判定およびエラーコード別の原因解説",
        48,
        "status_analysis (独自付与)",
        "String",
        "tls_verify_status の値に基づき、正常/失敗の判定と、AWS公式エラーコードに対応する日本語の原因説明を表示します。",
        "なし"
    ),
    (
        "conn_trace_id",
        "AWS項目 #12: conn_trace_id",
        "接続トレースID (conn_trace_id)\n【接続追跡用ユニークID】\n各接続を一意に識別するID(ALBアクセスログとの紐付け用外部キー)",
        38,
        "conn_trace_id (位置: 12)",
        "String",
        "各接続を識別するための一意の不透明IDです。接続確立後のリクエストのALBアクセスログにも同じIDが記録され、両ログの紐付けに使用できます。",
        "生成されなかった場合は -"
    ),
    (
        "tls_keyexchange",
        "AWS項目 #13: tls_keyexchange",
        "TLS鍵交換方式 (tls_keyexchange)\n【TLS/PQ-TLS鍵交換アルゴリズム】\nハンドシェイクで使用された鍵交換方式(X25519/ECDHE等)",
        26,
        "tls_keyexchange (位置: 13)",
        "String",
        "[HTTPSリスナー] TLSまたはポスト量子TLS(PQ-TLS)のハンドシェイクで使用された鍵交換アルゴリズムです。",
        "非SSL/TLSリクエスト、または12項目形式のログで出力されていない場合は -"
    ),
    (
        "elb",
        "AWS項目 #14: elb",
        "ロードバランサーID (elb)\n【対象ALBリソース識別子】\nリクエストを処理したALBのリソースID(ログ行またはファイル名から取得)",
        38,
        "elb (位置: 14 / ファイル名補完)",
        "String",
        "ロードバランサーのリソースID（例: app/my-alb/1234567890abcdef）です。12項目形式ログの場合はS3ファイル名から自動補完します。",
        "ログ行・ファイル名の双方から特定できない場合は -"
    ),
    (
        "ip_address",
        "AWS項目 #15: ip_address",
        "ALBノードIP (ip_address)\n【処理ALBノードIPアドレス】\nリクエストを処理したALBノードのIPアドレス(ログ行またはファイル名から取得)",
        24,
        "ip_address (位置: 15 / ファイル名補完)",
        "String (IPv4/IPv6)",
        "リクエストを処理したロードバランサーノードのIPアドレス（内部ALBの場合はプライベートIP）です。12項目形式ログの場合はファイル名から自動補完します。",
        "ログ行・ファイル名の双方から特定できない場合は -"
    ),
]

# S3上のALB接続ログファイル名パターン:
# conn_log_<account>_elasticloadbalancing_<region>_app.<lb_name>.<lb_id>_<end_time>_<node_ip>_<random>.log.gz
FILENAME_RE = re.compile(
    r"(?:conn_log_)?(?P<account>\d+)_elasticloadbalancing_(?P<region>[a-z0-9-]+)_"
    r"(?P<elb>app\.[^_]+)_(?P<endtime>\d{8}T\d{4}Z)_(?P<node_ip>[0-9a-fA-F.:]+)_(?P<rand>[a-zA-Z0-9]+)\.log\.gz$",
    re.IGNORECASE,
)

# 1行トークナイザ: ダブルクォートで囲まれた文字列(内部空白や\"を含む可能性あり)または非空白トークンを抽出
TOKEN_RE = re.compile(r'"((?:\\.|[^"\\])*)"|(\S+)')

# RHEL 9.8 標準の Python 3.9 でも任意の小数秒桁数 (0～9桁) の ISO 8601 を確実に解析する正規表現
ISO_TS_RE = re.compile(
    r"^(?P<Y>\d{4})-(?P<m>\d{2})-(?P<d>\d{2})[T\s]"
    r"(?P<H>\d{2}):(?P<M>\d{2}):(?P<S>\d{2})"
    r"(?:\.(?P<frac>\d+))?"
    r"(?P<tz>Z|[+-]\d{2}:?\d{2})?$"
)


def tokenize_conn_log_line(line: str):
    tokens = []
    for match in TOKEN_RE.finditer(line):
        quoted, unquoted = match.groups()
        if quoted is not None:
            tokens.append(quoted.replace('\\"', '"'))
        else:
            tokens.append(unquoted)
    return tokens


def parse_iso_timestamp(ts_str: str):
    if not ts_str or ts_str == "-":
        return None
    s = ts_str.strip()
    m = ISO_TS_RE.match(s)
    if m:
        try:
            frac = m.group("frac") or ""
            micro = int((frac + "000000")[:6])
            tz_str = m.group("tz")
            tz_info = timezone.utc
            if tz_str and tz_str != "Z":
                sign = -1 if tz_str.startswith("-") else 1
                clean_tz = tz_str[1:].replace(":", "")
                tz_h = int(clean_tz[:2])
                tz_m = int(clean_tz[2:4]) if len(clean_tz) >= 4 else 0
                tz_info = timezone(timedelta( minutes=sign * (tz_h * 60 + tz_m) ))
            return datetime(
                int(m.group("Y")),
                int(m.group("m")),
                int(m.group("d")),
                int(m.group("H")),
                int(m.group("M")),
                int(m.group("S")),
                micro,
                tzinfo=tz_info,
            )
        except Exception:
            return None
    try:
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def format_jst(dt: datetime) -> str:
    if dt is None:
        return "-"
    jst_dt = dt.astimezone(JST)
    if jst_dt.microsecond:
        return jst_dt.strftime("%Y-%m-%d %H:%M:%S.") + f"{jst_dt.microsecond // 1000:03d}"
    return jst_dt.strftime("%Y-%m-%d %H:%M:%S")


def format_cert_validity_jst(validity_str: str) -> str:
    if not validity_str or validity_str == "-":
        return "-"
    # Expected format: NotBefore=2023-09-21T22:43:21Z;NotAfter=2026-06-17T22:43:21Z
    nb_match = re.search(r"NotBefore=([^;,\s]+)", validity_str)
    na_match = re.search(r"NotAfter=([^;,\s]+)", validity_str)
    if not nb_match and not na_match:
        return validity_str
    nb_dt = parse_iso_timestamp(nb_match.group(1)) if nb_match else None
    na_dt = parse_iso_timestamp(na_match.group(1)) if na_match else None
    nb_str = nb_dt.astimezone(JST).strftime("%Y-%m-%d %H:%M:%S") if nb_dt else (nb_match.group(1) if nb_match else "?")
    na_str = na_dt.astimezone(JST).strftime("%Y-%m-%d %H:%M:%S") if na_dt else (na_match.group(1) if na_match else "?")
    return f"{nb_str} ～ {na_str} (JST)"


def analyze_status(tls_verify_status: str, tls_protocol: str) -> tuple:
    """
    Returns (category, error_code, status_analysis_text)
    category: 'SUCCESS', 'FAILED', 'NON_TLS'
    """
    if not tls_verify_status or tls_verify_status == "-":
        if not tls_protocol or tls_protocol == "-":
            return ("NON_TLS", "-", "非TLS接続 / 未検証 (-)")
        return ("NON_TLS", "-", "ステータス未記録 (-)")
    if tls_verify_status == "Success":
        return ("SUCCESS", "-", "正常確立 (Success)")
    if tls_verify_status.startswith("Failed"):
        parts = tls_verify_status.split(":", 1)
        err_code = parts[1].strip() if len(parts) > 1 and parts[1].strip() else "UnknownError"
        if err_code in AWS_ERROR_DESCRIPTIONS:
            short_ja, detail_ja = AWS_ERROR_DESCRIPTIONS[err_code]
            return ("FAILED", err_code, f"失敗: {short_ja} ({err_code}) - {detail_ja}")
        return ("FAILED", err_code, f"失敗: 接続/検証エラー ({err_code})")
    return ("NON_TLS", tls_verify_status, f"その他ステータス ({tls_verify_status})")


def extract_metadata_from_filename(filename: str) -> tuple:
    m = FILENAME_RE.search(filename)
    if not m:
        return ("-", "-")
    raw_elb = m.group("elb")
    # app.my-loadbalancer.1234567890abcdef -> app/my-loadbalancer/1234567890abcdef
    elb_parts = raw_elb.split(".")
    if len(elb_parts) >= 3 and elb_parts[0] == "app":
        elb_id = f"app/{'.'.join(elb_parts[1:-1])}/{elb_parts[-1]}"
    else:
        elb_id = raw_elb
    node_ip = m.group("node_ip")
    return (elb_id, node_ip)


def parse_single_line(line: str, rel_file: str, line_no: int, default_elb: str, default_node_ip: str):
    s = line.strip()
    if not s or s.startswith("#"):
        return None

    # JSON形式の接続ログにも対応
    if s.startswith("{") and s.endswith("}"):
        try:
            obj = json.loads(s)
            ts_str = str(obj.get("timestamp") or obj.get("time") or "-")
            client_ip = str(obj.get("client_ip") or "-")
            client_port = str(obj.get("client_port") or "-")
            listener_port = str(obj.get("listener_port") or "-")
            tls_protocol = str(obj.get("tls_protocol") or "-")
            tls_cipher = str(obj.get("tls_cipher") or "-")
            tls_handshake_latency = str(obj.get("tls_handshake_latency") or "-")
            leaf_subject = str(obj.get("leaf_client_cert_subject") or "-")
            leaf_validity = str(obj.get("leaf_client_cert_validity") or "-")
            leaf_serial = str(obj.get("leaf_client_cert_serial_number") or "-")
            tls_verify_status = str(obj.get("tls_verify_status") or "-")
            conn_trace_id = str(obj.get("conn_trace_id") or "-")
            tls_keyexchange = str(obj.get("tls_keyexchange") or "-")
            elb = str(obj.get("elb") or default_elb or "-")
            ip_address = str(obj.get("ip_address") or default_node_ip or "-")
        except Exception:
            return {"_malformed": True, "raw": s, "source_file": f"{rel_file}:{line_no}"}
    else:
        tokens = tokenize_conn_log_line(s)
        # ALB connection_log は標準12項目（拡張最大15項目以上）
        if len(tokens) < 11:
            return {"_malformed": True, "raw": s, "source_file": f"{rel_file}:{line_no}"}

        ts_str = tokens[0]
        client_ip = tokens[1]
        client_port = tokens[2]
        listener_port = tokens[3]
        tls_protocol = tokens[4]
        tls_cipher = tokens[5]
        tls_handshake_latency = tokens[6]
        leaf_subject = tokens[7] if tokens[7] != "" else "-"
        leaf_validity = tokens[8]
        leaf_serial = tokens[9]
        tls_verify_status = tokens[10]
        conn_trace_id = tokens[11] if len(tokens) >= 12 else "-"
        tls_keyexchange = tokens[12] if len(tokens) >= 13 else "-"
        elb = tokens[13] if len(tokens) >= 14 and tokens[13] != "-" else default_elb
        ip_address = tokens[14] if len(tokens) >= 15 and tokens[14] != "-" else default_node_ip

    dt = parse_iso_timestamp(ts_str)
    ts_jst = format_jst(dt)
    cert_val_jst = format_cert_validity_jst(leaf_validity)
    category, err_code, status_analysis = analyze_status(tls_verify_status, tls_protocol)

    # 数値変換 (Excel上で数値として扱えるように保持)
    try:
        client_port_val = int(client_port)
    except Exception:
        client_port_val = client_port

    try:
        listener_port_val = int(listener_port)
    except Exception:
        listener_port_val = listener_port

    latency_float = None
    if tls_handshake_latency != "-":
        try:
            latency_float = float(tls_handshake_latency)
        except Exception:
            latency_float = None

    return {
        "_malformed": False,
        "_dt": dt,
        "_sort_key": dt.timestamp() if dt else float("inf"),
        "_category": category,
        "_err_code": err_code,
        "_latency_float": latency_float,
        "_rel_file": rel_file,
        "source_file": f"{rel_file} (L{line_no})",
        "timestamp": ts_str,
        "timestamp_jst": ts_jst,
        "client_ip": client_ip,
        "client_port": client_port_val,
        "listener_port": listener_port_val,
        "tls_protocol": tls_protocol,
        "tls_cipher": tls_cipher,
        "tls_handshake_latency": latency_float if latency_float is not None else tls_handshake_latency,
        "leaf_client_cert_subject": leaf_subject,
        "leaf_client_cert_validity": leaf_validity,
        "cert_validity_jst": cert_val_jst,
        "leaf_client_cert_serial_number": leaf_serial,
        "tls_verify_status": tls_verify_status,
        "status_analysis": status_analysis,
        "conn_trace_id": conn_trace_id,
        "tls_keyexchange": tls_keyexchange,
        "elb": elb,
        "ip_address": ip_address,
    }


# ==============================================================================
# 2. OpenXML (.xlsx) ワークブック生成エンジン (標準ライブラリのみ・スタイル完備)
# ==============================================================================
# XML 1.0 で禁止されている制御文字を除去
ILLEGAL_XML_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def clean_xml_text(val) -> str:
    if val is None:
        return ""
    s = ILLEGAL_XML_CHARS_RE.sub("", str(val))
    return xml_escape(s)


def col_letter(col_idx_1based: int) -> str:
    res = []
    n = col_idx_1based
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res.append(chr(65 + rem))
    return "".join(reversed(res))


# スタイルID定義 (xl/styles.xml の cellXfs インデックスに対応)
STYLE_DEFAULT = 0          # 標準セル (左揃え・細罫線)
STYLE_TITLE = 1            # シートタイトルバナー (濃紺背景・白太字 14pt)
STYLE_SUBTITLE = 2         # サブタイトル・メタ情報バー (薄青背景・10pt)
STYLE_CAT_HEADER = 3       # カテゴリ・AWS項目番号ヘッダ (濃青灰背景・白太字・中央揃え)
STYLE_MAIN_HEADER = 4      # 各項目内容詳細ヘッダ (ネイビー背景・白太字・折返し・上揃え)
STYLE_SEC_HEADER = 5       # サマリーセクション見出し (濃紺左ボーダー風・薄青背景・太字 11pt)
STYLE_TBL_HEADER = 6       # サマリーテーブルヘッダ (ネイビー背景・白太字・中央揃え)
STYLE_ROW_EVEN = 7         # 通常行(白背景)・左揃え
STYLE_ROW_ODD = 8          # 通常行(ストライプ薄青背景)・左揃え
STYLE_ROW_EVEN_CENTER = 9  # 通常行(白背景)・中央揃え
STYLE_ROW_ODD_CENTER = 10  # 通常行(ストライプ薄青背景)・中央揃え
STYLE_ROW_EVEN_NUM = 11    # 通常行(白背景)・整数右揃え
STYLE_ROW_ODD_NUM = 12     # 通常行(ストライプ薄青背景)・整数右揃え
STYLE_ROW_EVEN_FLOAT = 13  # 通常行(白背景)・小数(0.000)右揃え
STYLE_ROW_ODD_FLOAT = 14   # 通常行(ストライプ薄青背景)・小数(0.000)右揃え
STYLE_STATUS_OK = 15       # Successステータスセル (薄緑背景・濃緑太字・中央揃え)
STYLE_ROW_FAIL = 16        # 失敗行(薄赤背景)・左揃え
STYLE_ROW_FAIL_CENTER = 17 # 失敗行(薄赤背景)・中央揃え
STYLE_ROW_FAIL_NUM = 18    # 失敗行(薄赤背景)・整数右揃え
STYLE_ROW_FAIL_FLOAT = 19  # 失敗行(薄赤背景)・小数右揃え
STYLE_STATUS_FAIL = 20     # Failedステータスセル (赤背景強調・濃赤太字)
STYLE_ROW_WARN = 21        # 非TLS/未検証行(薄黄背景)・左揃え
STYLE_ROW_WARN_CENTER = 22 # 非TLS/未検証行(薄黄背景)・中央揃え
STYLE_ROW_WARN_NUM = 23    # 非TLS/未検証行(薄黄背景)・整数右揃え
STYLE_KPI_LABEL = 24       # KPIラベル (薄灰青背景・太字)
STYLE_KPI_VALUE = 25       # KPI値 (白背景・太字・右/中央揃え)
STYLE_PERCENT = 26         # パーセント表示 (0.00%)


def build_styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="0.000"/>
    <numFmt numFmtId="165" formatCode="0.00%"/>
  </numFmts>
  <fonts count="8">
    <font><sz val="10"/><color rgb="FF1F2937"/><name val="Meiryo UI"/><family val="2"/></font>
    <font><b/><sz val="13"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="2"/></font>
    <font><b/><sz val="9.5"/><color rgb="FF1E3A8A"/><name val="Meiryo UI"/><family val="2"/></font>
    <font><b/><sz val="9"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="2"/></font>
    <font><b/><sz val="9.5"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF0F172A"/><name val="Meiryo UI"/><family val="2"/></font>
    <font><b/><sz val="10"/><color rgb="FF14532D"/><name val="Meiryo UI"/><family val="2"/></font>
    <font><b/><sz val="10"/><color rgb="FF991B1B"/><name val="Meiryo UI"/><family val="2"/></font>
  </fonts>
  <fills count="12">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E3A8A"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEFF6FF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF334155"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E40AF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFECACA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF9C3"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FFCBD5E1"/></top>
      <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF94A3B8"/></left>
      <right style="thin"><color rgb="FF94A3B8"/></right>
      <top style="medium"><color rgb="FF1E3A8A"/></top>
      <bottom style="medium"><color rgb="FF1E3A8A"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="27">
    <!-- 0: STYLE_DEFAULT -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment vertical="center" wrapText="0"/>
    </xf>
    <!-- 1: STYLE_TITLE -->
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 2: STYLE_SUBTITLE -->
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 3: STYLE_CAT_HEADER -->
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <!-- 4: STYLE_MAIN_HEADER -->
    <xf numFmtId="0" fontId="4" fillId="5" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="top" wrapText="1"/>
    </xf>
    <!-- 5: STYLE_SEC_HEADER -->
    <xf numFmtId="0" fontId="5" fillId="6" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 6: STYLE_TBL_HEADER -->
    <xf numFmtId="0" fontId="4" fillId="5" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <!-- 7: STYLE_ROW_EVEN -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 8: STYLE_ROW_ODD -->
    <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 9: STYLE_ROW_EVEN_CENTER -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 10: STYLE_ROW_ODD_CENTER -->
    <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 11: STYLE_ROW_EVEN_NUM -->
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 12: STYLE_ROW_ODD_NUM -->
    <xf numFmtId="1" fontId="0" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 13: STYLE_ROW_EVEN_FLOAT -->
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 14: STYLE_ROW_ODD_FLOAT -->
    <xf numFmtId="164" fontId="0" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 15: STYLE_STATUS_OK -->
    <xf numFmtId="0" fontId="6" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 16: STYLE_ROW_FAIL -->
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 17: STYLE_ROW_FAIL_CENTER -->
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 18: STYLE_ROW_FAIL_NUM -->
    <xf numFmtId="1" fontId="0" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 19: STYLE_ROW_FAIL_FLOAT -->
    <xf numFmtId="164" fontId="0" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 20: STYLE_STATUS_FAIL -->
    <xf numFmtId="0" fontId="7" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 21: STYLE_ROW_WARN -->
    <xf numFmtId="0" fontId="0" fillId="11" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 22: STYLE_ROW_WARN_CENTER -->
    <xf numFmtId="0" fontId="0" fillId="11" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 23: STYLE_ROW_WARN_NUM -->
    <xf numFmtId="1" fontId="0" fillId="11" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 24: STYLE_KPI_LABEL -->
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 25: STYLE_KPI_VALUE -->
    <xf numFmtId="0" fontId="5" fillId="0" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 26: STYLE_PERCENT -->
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""


class SharedStringsTable:
    def __init__(self):
        self._map = {}
        self._list = []
        self.total_refs = 0

    def get_index(self, text: str) -> int:
        self.total_refs += 1
        s = str(text) if text is not None else ""
        idx = self._map.get(s)
        if idx is None:
            idx = len(self._list)
            self._map[s] = idx
            self._list.append(s)
        return idx

    def to_xml(self) -> str:
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{self.total_refs}" uniqueCount="{len(self._list)}">'
        ]
        for s in self._list:
            cleaned = clean_xml_text(s)
            if s.startswith(" ") or s.endswith(" ") or "\n" in s or "\t" in s:
                parts.append(f'<si><t xml:space="preserve">{cleaned}</t></si>')
            else:
                parts.append(f"<si><t>{cleaned}</t></si>")
        parts.append("</sst>")
        return "\n".join(parts)


class WorksheetBuilder:
    def __init__(self, name: str, sst: SharedStringsTable, col_widths: list, freeze_row: int = 0, auto_filter_ref: str = None):
        self.name = name
        self.sst = sst
        self.col_widths = col_widths
        self.freeze_row = freeze_row
        self.auto_filter_ref = auto_filter_ref
        self.rows_xml = []
        self.merges = []
        self.current_row = 0

    def add_row(self, cells: list, height: float = None):
        """
        cells: list of (value, style_id) or value (uses STYLE_DEFAULT)
        """
        self.current_row += 1
        r_idx = self.current_row
        ht_attr = f' ht="{height}" customHeight="1"' if height is not None else ""
        cell_parts = []
        for c_idx, item in enumerate(cells, start=1):
            if isinstance(item, tuple):
                val, style_id = item
            else:
                val, style_id = item, STYLE_DEFAULT

            if val is None or val == "":
                cell_ref = f"{col_letter(c_idx)}{r_idx}"
                cell_parts.append(f'<c r="{cell_ref}" s="{style_id}"/>')
                continue

            cell_ref = f"{col_letter(c_idx)}{r_idx}"
            if isinstance(val, bool):
                s_idx = self.sst.get_index(str(val))
                cell_parts.append(f'<c r="{cell_ref}" s="{style_id}" t="s"><v>{s_idx}</v></c>')
            elif isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                cell_parts.append(f'<c r="{cell_ref}" s="{style_id}"><v>{val}</v></c>')
            else:
                s_idx = self.sst.get_index(str(val))
                cell_parts.append(f'<c r="{cell_ref}" s="{style_id}" t="s"><v>{s_idx}</v></c>')

        self.rows_xml.append(f'<row r="{r_idx}"{ht_attr}>{"".join(cell_parts)}</row>')
        return r_idx

    def merge_cells(self, start_row: int, start_col: int, end_row: int, end_col: int):
        ref = f"{col_letter(start_col)}{start_row}:{col_letter(end_col)}{end_row}"
        self.merges.append(ref)

    def to_xml(self) -> str:
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        ]
        # 1. sheetViews (Freeze Panes)
        if self.freeze_row and self.freeze_row > 0:
            top_left = f"A{self.freeze_row + 1}"
            parts.append(
                f'<sheetViews><sheetView workbookViewId="0">'
                f'<pane ySplit="{self.freeze_row}" topLeftCell="{top_left}" activePane="bottomLeft" state="frozen"/>'
                f'<selection pane="bottomLeft" activeCell="{top_left}" sqref="{top_left}"/>'
                f'</sheetView></sheetViews>'
            )
        else:
            parts.append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')

        # 2. sheetFormatPr
        parts.append('<sheetFormatPr defaultRowHeight="19"/>')

        # 3. cols
        if self.col_widths:
            parts.append("<cols>")
            for idx, w in enumerate(self.col_widths, start=1):
                parts.append(f'<col min="{idx}" max="{idx}" width="{w}" customWidth="1"/>')
            parts.append("</cols>")

        # 4. sheetData
        parts.append("<sheetData>")
        parts.extend(self.rows_xml)
        parts.append("</sheetData>")

        # 5. mergeCells (OpenXML仕様上、sheetDataの直後・autoFilterの前に配置)
        if self.merges:
            parts.append(f'<mergeCells count="{len(self.merges)}">')
            for m_ref in self.merges:
                parts.append(f'<mergeCell ref="{m_ref}"/>')
            parts.append("</mergeCells>")

        # 6. autoFilter
        if self.auto_filter_ref:
            parts.append(f'<autoFilter ref="{self.auto_filter_ref}"/>')

        # 7. pageMargins
        parts.append('<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        parts.append("</worksheet>")
        return "\n".join(parts)


# ==============================================================================
# 3. シート構築処理
# ==============================================================================
def build_detail_sheet(
    sheet_name: str,
    title_text: str,
    subtitle_text: str,
    records: list,
    sst: SharedStringsTable,
    empty_message: str = None,
) -> WorksheetBuilder:
    col_widths = [col[3] for col in DETAIL_COLUMNS]
    num_cols = len(DETAIL_COLUMNS)
    last_col_letter = col_letter(num_cols)
    last_data_row = 4 + max(len(records), 1)
    auto_filter_ref = f"A4:{last_col_letter}{last_data_row}"

    ws = WorksheetBuilder(
        name=sheet_name,
        sst=sst,
        col_widths=col_widths,
        freeze_row=4,
        auto_filter_ref=auto_filter_ref,
    )

    # Row 1: タイトルバナー
    r1_cells = [(title_text, STYLE_TITLE)] + [("", STYLE_TITLE) for _ in range(num_cols - 1)]
    ws.add_row(r1_cells, height=28)
    ws.merge_cells(1, 1, 1, num_cols)

    # Row 2: サマリー情報バー
    r2_cells = [(subtitle_text, STYLE_SUBTITLE)] + [("", STYLE_SUBTITLE) for _ in range(num_cols - 1)]
    ws.add_row(r2_cells, height=22)
    ws.merge_cells(2, 1, 2, num_cols)

    # Row 3: AWS公式項目番号・カテゴリヘッダ
    r3_cells = [(col[1], STYLE_CAT_HEADER) for col in DETAIL_COLUMNS]
    ws.add_row(r3_cells, height=22)

    # Row 4: 各出力項目の内容・詳細説明ヘッダ (3行折り返し表示)
    r4_cells = [(col[2], STYLE_MAIN_HEADER) for col in DETAIL_COLUMNS]
    ws.add_row(r4_cells, height=54)

    # データ行が空の場合
    if not records:
        msg = empty_message or "該当する接続ログレコードはありません。"
        empty_cells = [(msg, STYLE_ROW_EVEN)] + [("", STYLE_ROW_EVEN) for _ in range(num_cols - 1)]
        ws.add_row(empty_cells, height=24)
        ws.merge_cells(5, 1, 5, num_cols)
        return ws

    # Row 5+: 1行ずつわかりやすく確認できる接続ログレコード
    for idx, rec in enumerate(records, start=1):
        cat = rec["_category"]
        is_odd = (idx % 2 == 1)

        if cat == "FAILED":
            s_text = STYLE_ROW_FAIL
            s_center = STYLE_ROW_FAIL_CENTER
            s_num = STYLE_ROW_FAIL_NUM
            s_float = STYLE_ROW_FAIL_FLOAT
            s_status = STYLE_STATUS_FAIL
        elif cat == "NON_TLS":
            s_text = STYLE_ROW_WARN
            s_center = STYLE_ROW_WARN_CENTER
            s_num = STYLE_ROW_WARN_NUM
            s_float = STYLE_ROW_WARN_CENTER
            s_status = STYLE_ROW_WARN_CENTER
        else:
            s_text = STYLE_ROW_ODD if is_odd else STYLE_ROW_EVEN
            s_center = STYLE_ROW_ODD_CENTER if is_odd else STYLE_ROW_EVEN_CENTER
            s_num = STYLE_ROW_ODD_NUM if is_odd else STYLE_ROW_EVEN_NUM
            s_float = STYLE_ROW_ODD_FLOAT if is_odd else STYLE_ROW_EVEN_FLOAT
            s_status = STYLE_STATUS_OK

        lat_val = rec["tls_handshake_latency"]
        lat_style = s_float if isinstance(lat_val, float) else s_center

        c_port = rec["client_port"]
        c_port_style = s_num if isinstance(c_port, int) else s_center

        l_port = rec["listener_port"]
        l_port_style = s_num if isinstance(l_port, int) else s_center

        row_cells = [
            (idx, s_num),                                      # 1. No.
            (rec["source_file"], s_text),                      # 2. source_file
            (rec["timestamp"], s_center),                      # 3. timestamp (UTC)
            (rec["timestamp_jst"], s_center),                  # 4. timestamp (JST)
            (rec["client_ip"], s_center),                      # 5. client_ip
            (c_port, c_port_style),                            # 6. client_port
            (l_port, l_port_style),                            # 7. listener_port
            (rec["tls_protocol"], s_center),                   # 8. tls_protocol
            (rec["tls_cipher"], s_text),                       # 9. tls_cipher
            (lat_val, lat_style),                              # 10. tls_handshake_latency
            (rec["leaf_client_cert_subject"], s_text),         # 11. leaf_client_cert_subject
            (rec["leaf_client_cert_validity"], s_text),        # 12. leaf_client_cert_validity
            (rec["cert_validity_jst"], s_text),                # 13. cert_validity_jst
            (rec["leaf_client_cert_serial_number"], s_center), # 14. leaf_client_cert_serial_number
            (rec["tls_verify_status"], s_status),              # 15. tls_verify_status
            (rec["status_analysis"], s_status if cat == "FAILED" else s_text),  # 16. status_analysis
            (rec["conn_trace_id"], s_text),                    # 17. conn_trace_id
            (rec["tls_keyexchange"], s_center),                # 18. tls_keyexchange
            (rec["elb"], s_text),                              # 19. elb
            (rec["ip_address"], s_center),                     # 20. ip_address
        ]
        ws.add_row(row_cells, height=20)

    return ws


def percentile(sorted_vals: list, p: float) -> float:
    if not sorted_vals:
        return 0.0
    k = (len(sorted_vals) - 1) * p
    f = math.floor(k)
    c = math.ceil(k)
    if f == c:
        return sorted_vals[int(k)]
    return sorted_vals[f] * (c - k) + sorted_vals[c] * (k - f)


def build_summary_sheet(
    input_dir: str,
    file_stats: list,
    records: list,
    malformed_rows: list,
    sst: SharedStringsTable,
) -> WorksheetBuilder:
    col_widths = [28, 38, 18, 18, 18, 18, 22, 48]
    num_cols = len(col_widths)
    ws = WorksheetBuilder(
        name="分析サマリーレポート",
        sst=sst,
        col_widths=col_widths,
        freeze_row=2,
    )

    now_jst = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    total_count = len(records)
    success_count = sum(1 for r in records if r["_category"] == "SUCCESS")
    failed_count = sum(1 for r in records if r["_category"] == "FAILED")
    non_tls_count = sum(1 for r in records if r["_category"] == "NON_TLS")
    success_rate = (success_count / total_count) if total_count > 0 else 0.0
    fail_rate = (failed_count / total_count) if total_count > 0 else 0.0

    valid_dts = [r["_dt"] for r in records if r["_dt"] is not None]
    if valid_dts:
        min_dt = min(valid_dts)
        max_dt = max(valid_dts)
        period_utc = f"{min_dt.strftime('%Y-%m-%d %H:%M:%S')} UTC ～ {max_dt.strftime('%Y-%m-%d %H:%M:%S')} UTC"
        period_jst = f"{format_jst(min_dt)} JST ～ {format_jst(max_dt)} JST"
    else:
        period_utc = "-"
        period_jst = "-"

    latencies = sorted(r["_latency_float"] for r in records if r["_latency_float"] is not None)
    if latencies:
        lat_min = min(latencies)
        lat_avg = sum(latencies) / len(latencies)
        lat_med = percentile(latencies, 0.50)
        lat_p95 = percentile(latencies, 0.95)
        lat_max = max(latencies)
        lat_summary = (
            f"最小: {lat_min:.3f}秒 / 平均: {lat_avg:.3f}秒 / 中央値: {lat_med:.3f}秒 / "
            f"95%タイル: {lat_p95:.3f}秒 / 最大: {lat_max:.3f}秒 (計測対象: {len(latencies):,}件)"
        )
    else:
        lat_summary = "計測対象データなし (-)"

    # Row 1: Title
    r1 = [("AWS ALB 接続ログ (Connection Log) 分析サマリーレポート", STYLE_TITLE)] + [("", STYLE_TITLE) for _ in range(num_cols - 1)]
    ws.add_row(r1, height=28)
    ws.merge_cells(1, 1, 1, num_cols)

    # Row 2: Subtitle
    r2 = [(f"レポート生成日時: {now_jst} | 対象ディレクトリ: {input_dir} | 読込gzファイル数: {len(file_stats)}件 | 総接続ログ数: {total_count:,}件", STYLE_SUBTITLE)] + [("", STYLE_SUBTITLE) for _ in range(num_cols - 1)]
    ws.add_row(r2, height=22)
    ws.merge_cells(2, 1, 2, num_cols)

    ws.add_row([("", STYLE_DEFAULT)], height=10)

    def add_section_header(title: str):
        r = ws.add_row([(title, STYLE_SEC_HEADER)] + [("", STYLE_SEC_HEADER) for _ in range(num_cols - 1)], height=24)
        ws.merge_cells(r, 1, r, num_cols)

    # --------------------------------------------------------------------------
    # セクション 1: 全体集計サマリー (KPI)
    # --------------------------------------------------------------------------
    add_section_header("1. 全体集計サマリー (Overview & KPIs)")
    kpi_rows = [
        ("解析対象ディレクトリ", str(input_dir), "処理対象 .gz ファイル数", f"{len(file_stats):,} ファイル"),
        ("総接続ログ件数 (Total)", f"{total_count:,} 件", "フォーマット不正・スキップ行数", f"{len(malformed_rows):,} 行"),
        ("正常接続件数 (Success)", f"{success_count:,} 件 ({success_rate * 100:.2f}%)", "失敗接続件数 (Failed:*)", f"{failed_count:,} 件 ({fail_rate * 100:.2f}%)"),
        ("非TLS/ステータス未設定 (-)", f"{non_tls_count:,} 件", "ユニーククライアントIP数", f"{len({r['client_ip'] for r in records}):,} IP"),
        ("ログ記録期間 (UTC)", period_utc, "ログ記録期間 (JST 日本時間)", period_jst),
        ("TLSハンドシェイク所要時間統計", lat_summary, "ユニーククライアント証明書数", f"{len({r['leaf_client_cert_serial_number'] for r in records if r['leaf_client_cert_serial_number'] != '-'}):,} 枚"),
    ]
    for k1, v1, k2, v2 in kpi_rows:
        r = ws.add_row([
            (k1, STYLE_KPI_LABEL),
            (v1, STYLE_KPI_VALUE),
            ("", STYLE_KPI_VALUE),
            (k2, STYLE_KPI_LABEL),
            (v2, STYLE_KPI_VALUE),
            ("", STYLE_KPI_VALUE),
            ("", STYLE_KPI_VALUE),
            ("", STYLE_KPI_VALUE),
        ], height=21)
        ws.merge_cells(r, 2, r, 3)
        ws.merge_cells(r, 5, r, 8)

    ws.add_row([("", STYLE_DEFAULT)], height=12)

    # --------------------------------------------------------------------------
    # セクション 2: TLS接続検証ステータス (tls_verify_status) 別集計
    # --------------------------------------------------------------------------
    add_section_header("2. TLS接続検証ステータス・エラーコード別集計 (tls_verify_status Breakdown)")
    ws.add_row([
        ("TLS接続検証ステータス\n(tls_verify_status)", STYLE_TBL_HEADER),
        ("エラーコード\n(error_code)", STYLE_TBL_HEADER),
        ("件数\n(Count)", STYLE_TBL_HEADER),
        ("構成比\n(Share)", STYLE_TBL_HEADER),
        ("平均ハンドシェイク[秒]\n(Avg Latency)", STYLE_TBL_HEADER),
        ("最大ハンドシェイク[秒]\n(Max Latency)", STYLE_TBL_HEADER),
        ("影響クライアントIP数\n(Unique IPs)", STYLE_TBL_HEADER),
        ("判定・エラー原因の日本語解説\n(Description & Troubleshooting)", STYLE_TBL_HEADER),
    ], height=32)

    status_groups = defaultdict(list)
    for r in records:
        status_groups[r["tls_verify_status"]].append(r)

    for st_val, grp in sorted(status_groups.items(), key=lambda x: (-len(x[1]), x[0])):
        cnt = len(grp)
        share = cnt / total_count if total_count > 0 else 0.0
        grp_lats = [x["_latency_float"] for x in grp if x["_latency_float"] is not None]
        avg_lat = round(sum(grp_lats) / len(grp_lats), 3) if grp_lats else "-"
        max_lat = round(max(grp_lats), 3) if grp_lats else "-"
        uniq_ips = len({x["client_ip"] for x in grp})
        cat = grp[0]["_category"]
        err_code = grp[0]["_err_code"]
        desc = grp[0]["status_analysis"]

        row_style = STYLE_ROW_FAIL if cat == "FAILED" else (STYLE_ROW_WARN if cat == "NON_TLS" else STYLE_ROW_EVEN)
        num_style = STYLE_ROW_FAIL_NUM if cat == "FAILED" else (STYLE_ROW_WARN_NUM if cat == "NON_TLS" else STYLE_ROW_EVEN_NUM)
        st_cell_style = STYLE_STATUS_FAIL if cat == "FAILED" else (STYLE_STATUS_OK if cat == "SUCCESS" else STYLE_ROW_WARN_CENTER)

        ws.add_row([
            (st_val, st_cell_style),
            (err_code, row_style),
            (cnt, num_style),
            (share, STYLE_PERCENT),
            (avg_lat, STYLE_ROW_EVEN_FLOAT if isinstance(avg_lat, float) else STYLE_ROW_EVEN_CENTER),
            (max_lat, STYLE_ROW_EVEN_FLOAT if isinstance(max_lat, float) else STYLE_ROW_EVEN_CENTER),
            (uniq_ips, num_style),
            (desc, row_style),
        ], height=21)

    ws.add_row([("", STYLE_DEFAULT)], height=12)

    # --------------------------------------------------------------------------
    # セクション 3: TLSプロトコル × 暗号スイート × 鍵交換方式別集計
    # --------------------------------------------------------------------------
    add_section_header("3. TLSプロトコル・暗号スイート・鍵交換方式別集計 (TLS Protocol & Cipher Suite Breakdown)")
    ws.add_row([
        ("TLSプロトコル\n(tls_protocol)", STYLE_TBL_HEADER),
        ("TLS暗号スイート\n(tls_cipher)", STYLE_TBL_HEADER),
        ("TLS鍵交換方式\n(tls_keyexchange)", STYLE_TBL_HEADER),
        ("総接続件数\n(Total)", STYLE_TBL_HEADER),
        ("構成比\n(Share)", STYLE_TBL_HEADER),
        ("成功件数\n(Success)", STYLE_TBL_HEADER),
        ("失敗件数\n(Failed)", STYLE_TBL_HEADER),
        ("ハンドシェイク時間 (平均 / 最大 [秒])", STYLE_TBL_HEADER),
    ], height=32)

    tls_groups = defaultdict(list)
    for r in records:
        key = (r["tls_protocol"], r["tls_cipher"], r["tls_keyexchange"])
        tls_groups[key].append(r)

    for (proto, cipher, kex), grp in sorted(tls_groups.items(), key=lambda x: (-len(x[1]), x[0])):
        cnt = len(grp)
        share = cnt / total_count if total_count > 0 else 0.0
        ok_cnt = sum(1 for x in grp if x["_category"] == "SUCCESS")
        ng_cnt = sum(1 for x in grp if x["_category"] == "FAILED")
        grp_lats = [x["_latency_float"] for x in grp if x["_latency_float"] is not None]
        if grp_lats:
            lat_info = f"平均: {sum(grp_lats)/len(grp_lats):.3f}秒 / 最大: {max(grp_lats):.3f}秒"
        else:
            lat_info = "-"

        ws.add_row([
            (proto, STYLE_ROW_EVEN_CENTER),
            (cipher, STYLE_ROW_EVEN),
            (kex, STYLE_ROW_EVEN_CENTER),
            (cnt, STYLE_ROW_EVEN_NUM),
            (share, STYLE_PERCENT),
            (ok_cnt, STYLE_ROW_EVEN_NUM),
            (ng_cnt, STYLE_ROW_FAIL_NUM if ng_cnt > 0 else STYLE_ROW_EVEN_NUM),
            (lat_info, STYLE_ROW_EVEN),
        ], height=20)

    ws.add_row([("", STYLE_DEFAULT)], height=12)

    # --------------------------------------------------------------------------
    # セクション 4: クライアントIPアドレス別 接続数・エラー集計 (上位50件)
    # --------------------------------------------------------------------------
    add_section_header("4. クライアントIPアドレス別 接続・エラー集計 (Top Client IPs)")
    ws.add_row([
        ("クライアントIP\n(client_ip)", STYLE_TBL_HEADER),
        ("使用クライアント証明書Subject (代表値)\n(leaf_client_cert_subject)", STYLE_TBL_HEADER),
        ("総接続数\n(Total)", STYLE_TBL_HEADER),
        ("成功数\n(Success)", STYLE_TBL_HEADER),
        ("失敗数\n(Failed)", STYLE_TBL_HEADER),
        ("失敗率\n(Fail Rate)", STYLE_TBL_HEADER),
        ("使用TLSプロトコル\n(Protocols)", STYLE_TBL_HEADER),
        ("発生した主なステータス・エラー内容\n(Statuses)", STYLE_TBL_HEADER),
    ], height=32)

    ip_groups = defaultdict(list)
    for r in records:
        ip_groups[r["client_ip"]].append(r)

    for cip, grp in sorted(ip_groups.items(), key=lambda x: (-len(x[1]), -sum(1 for r in x[1] if r["_category"] == "FAILED"), x[0]))[:50]:
        cnt = len(grp)
        ok_cnt = sum(1 for x in grp if x["_category"] == "SUCCESS")
        ng_cnt = sum(1 for x in grp if x["_category"] == "FAILED")
        ng_rate = ng_cnt / cnt if cnt > 0 else 0.0
        subjects = sorted({x["leaf_client_cert_subject"] for x in grp if x["leaf_client_cert_subject"] != "-"})
        subj_str = ", ".join(subjects[:2]) + (f" 他{len(subjects)-2}件" if len(subjects) > 2 else "") if subjects else "-"
        protos = ", ".join(sorted({x["tls_protocol"] for x in grp}))
        st_counter = Counter(x["tls_verify_status"] for x in grp)
        st_str = ", ".join(f"{k}({v})" for k, v in st_counter.most_common(3))

        row_s = STYLE_ROW_FAIL if ng_cnt > 0 else STYLE_ROW_EVEN
        num_s = STYLE_ROW_FAIL_NUM if ng_cnt > 0 else STYLE_ROW_EVEN_NUM

        ws.add_row([
            (cip, STYLE_ROW_FAIL_CENTER if ng_cnt > 0 else STYLE_ROW_EVEN_CENTER),
            (subj_str, row_s),
            (cnt, num_s),
            (ok_cnt, num_s),
            (ng_cnt, num_s),
            (ng_rate, STYLE_PERCENT),
            (protos, STYLE_ROW_EVEN_CENTER),
            (st_str, row_s),
        ], height=20)

    ws.add_row([("", STYLE_DEFAULT)], height=12)

    # --------------------------------------------------------------------------
    # セクション 5: mTLS クライアント証明書別集計
    # --------------------------------------------------------------------------
    add_section_header("5. mTLS クライアント証明書別集計 (Client Certificates Summary)")
    ws.add_row([
        ("証明書シリアル番号\n(leaf_client_cert_serial_number)", STYLE_TBL_HEADER),
        ("証明書サブジェクト (DN)\n(leaf_client_cert_subject)", STYLE_TBL_HEADER),
        ("総接続数\n(Total)", STYLE_TBL_HEADER),
        ("成功数\n(Success)", STYLE_TBL_HEADER),
        ("失敗数\n(Failed)", STYLE_TBL_HEADER),
        ("クライアントIP数\n(Unique IPs)", STYLE_TBL_HEADER),
        ("主な検証ステータス\n(tls_verify_status)", STYLE_TBL_HEADER),
        ("証明書有効期間 (JST)\n(cert_validity_jst)", STYLE_TBL_HEADER),
    ], height=32)

    cert_groups = defaultdict(list)
    for r in records:
        key = (r["leaf_client_cert_serial_number"], r["leaf_client_cert_subject"], r["cert_validity_jst"])
        cert_groups[key].append(r)

    for (serial, subj, val_jst), grp in sorted(cert_groups.items(), key=lambda x: (-len(x[1]), x[0])):
        cnt = len(grp)
        ok_cnt = sum(1 for x in grp if x["_category"] == "SUCCESS")
        ng_cnt = sum(1 for x in grp if x["_category"] == "FAILED")
        uniq_ips = len({x["client_ip"] for x in grp})
        st_counter = Counter(x["tls_verify_status"] for x in grp)
        st_str = ", ".join(f"{k}({v})" for k, v in st_counter.most_common(3))
        row_s = STYLE_ROW_FAIL if ng_cnt > 0 else STYLE_ROW_EVEN
        num_s = STYLE_ROW_FAIL_NUM if ng_cnt > 0 else STYLE_ROW_EVEN_NUM

        ws.add_row([
            (serial, STYLE_ROW_FAIL_CENTER if ng_cnt > 0 else STYLE_ROW_EVEN_CENTER),
            (subj, row_s),
            (cnt, num_s),
            (ok_cnt, num_s),
            (ng_cnt, num_s),
            (uniq_ips, num_s),
            (st_str, row_s),
            (val_jst, row_s),
        ], height=20)

    ws.add_row([("", STYLE_DEFAULT)], height=12)

    # --------------------------------------------------------------------------
    # セクション 6: リスナーポート & ALBノード別集計
    # --------------------------------------------------------------------------
    add_section_header("6. リスナーポート・ALBノード別集計 (Listener Port & ALB Node Summary)")
    ws.add_row([
        ("ロードバランサーID\n(elb)", STYLE_TBL_HEADER),
        ("ALBノードIP\n(ip_address)", STYLE_TBL_HEADER),
        ("リスナーポート\n(listener_port)", STYLE_TBL_HEADER),
        ("総接続数\n(Total)", STYLE_TBL_HEADER),
        ("構成比\n(Share)", STYLE_TBL_HEADER),
        ("成功数\n(Success)", STYLE_TBL_HEADER),
        ("失敗数\n(Failed)", STYLE_TBL_HEADER),
        ("非TLS/未設定数\n(Non-TLS)", STYLE_TBL_HEADER),
    ], height=32)

    node_groups = defaultdict(list)
    for r in records:
        node_groups[(r["elb"], r["ip_address"], str(r["listener_port"]))].append(r)

    for (elb_id, node_ip, lport), grp in sorted(node_groups.items(), key=lambda x: (-len(x[1]), x[0])):
        cnt = len(grp)
        share = cnt / total_count if total_count > 0 else 0.0
        ok_cnt = sum(1 for x in grp if x["_category"] == "SUCCESS")
        ng_cnt = sum(1 for x in grp if x["_category"] == "FAILED")
        nt_cnt = sum(1 for x in grp if x["_category"] == "NON_TLS")
        ws.add_row([
            (elb_id, STYLE_ROW_EVEN),
            (node_ip, STYLE_ROW_EVEN_CENTER),
            (int(lport) if lport.isdigit() else lport, STYLE_ROW_EVEN_CENTER),
            (cnt, STYLE_ROW_EVEN_NUM),
            (share, STYLE_PERCENT),
            (ok_cnt, STYLE_ROW_EVEN_NUM),
            (ng_cnt, STYLE_ROW_FAIL_NUM if ng_cnt > 0 else STYLE_ROW_EVEN_NUM),
            (nt_cnt, STYLE_ROW_EVEN_NUM),
        ], height=20)

    ws.add_row([("", STYLE_DEFAULT)], height=12)

    # --------------------------------------------------------------------------
    # セクション 7: 時間帯別（1時間単位）接続推移
    # --------------------------------------------------------------------------
    add_section_header("7. 時間帯別 接続推移 (Hourly Connection Trend)")
    ws.add_row([
        ("時間帯 (JST 日本時間)\n(Hour JST)", STYLE_TBL_HEADER),
        ("時間帯 (UTC)\n(Hour UTC)", STYLE_TBL_HEADER),
        ("総接続数\n(Total)", STYLE_TBL_HEADER),
        ("成功数\n(Success)", STYLE_TBL_HEADER),
        ("失敗数\n(Failed)", STYLE_TBL_HEADER),
        ("失敗率\n(Fail Rate)", STYLE_TBL_HEADER),
        ("平均ハンドシェイク[秒]\n(Avg Latency)", STYLE_TBL_HEADER),
        ("時間帯内の主なエラー\n(Top Errors)", STYLE_TBL_HEADER),
    ], height=32)

    hour_groups = defaultdict(list)
    for r in records:
        if r["_dt"] is not None:
            h_utc = r["_dt"].astimezone(timezone.utc).strftime("%Y-%m-%d %H:00 UTC")
            h_jst = r["_dt"].astimezone(JST).strftime("%Y-%m-%d %H:00 JST")
            hour_groups[(h_jst, h_utc)].append(r)
        else:
            hour_groups[("日時不明", "日時不明")].append(r)

    for (h_jst, h_utc), grp in sorted(hour_groups.items(), key=lambda x: x[0]):
        cnt = len(grp)
        ok_cnt = sum(1 for x in grp if x["_category"] == "SUCCESS")
        ng_cnt = sum(1 for x in grp if x["_category"] == "FAILED")
        ng_rate = ng_cnt / cnt if cnt > 0 else 0.0
        grp_lats = [x["_latency_float"] for x in grp if x["_latency_float"] is not None]
        avg_lat = round(sum(grp_lats) / len(grp_lats), 3) if grp_lats else "-"
        err_counter = Counter(x["_err_code"] for x in grp if x["_category"] == "FAILED")
        err_str = ", ".join(f"{k}({v})" for k, v in err_counter.most_common(3)) if err_counter else "エラーなし"

        ws.add_row([
            (h_jst, STYLE_ROW_EVEN_CENTER),
            (h_utc, STYLE_ROW_EVEN_CENTER),
            (cnt, STYLE_ROW_EVEN_NUM),
            (ok_cnt, STYLE_ROW_EVEN_NUM),
            (ng_cnt, STYLE_ROW_FAIL_NUM if ng_cnt > 0 else STYLE_ROW_EVEN_NUM),
            (ng_rate, STYLE_PERCENT),
            (avg_lat, STYLE_ROW_EVEN_FLOAT if isinstance(avg_lat, float) else STYLE_ROW_EVEN_CENTER),
            (err_str, STYLE_ROW_FAIL if ng_cnt > 0 else STYLE_ROW_EVEN),
        ], height=20)

    ws.add_row([("", STYLE_DEFAULT)], height=12)

    # --------------------------------------------------------------------------
    # セクション 8: 読込 .gz ファイル別集計
    # --------------------------------------------------------------------------
    add_section_header("8. 読込 .gz ファイル別集計 (Processed .gz Files Summary)")
    ws.add_row([
        ("No.", STYLE_TBL_HEADER),
        ("読込 .gz ファイル名 (相対パス)\n(File Name)", STYLE_TBL_HEADER),
        ("ファイルサイズ(Byte)\n(Size)", STYLE_TBL_HEADER),
        ("有効レコード数\n(Valid Rows)", STYLE_TBL_HEADER),
        ("成功数\n(Success)", STYLE_TBL_HEADER),
        ("失敗数\n(Failed)", STYLE_TBL_HEADER),
        ("読込ステータス\n(File Status)", STYLE_TBL_HEADER),
        ("ファイル内ログ期間 (JST)\n(Time Range JST)", STYLE_TBL_HEADER),
    ], height=32)

    for idx, fstat in enumerate(file_stats, start=1):
        is_err = fstat["status"] != "OK"
        ws.add_row([
            (idx, STYLE_ROW_EVEN_NUM),
            (fstat["rel_path"], STYLE_ROW_FAIL if is_err else STYLE_ROW_EVEN),
            (fstat["size_bytes"], STYLE_ROW_EVEN_NUM),
            (fstat["records"], STYLE_ROW_EVEN_NUM),
            (fstat["success"], STYLE_ROW_EVEN_NUM),
            (fstat["failed"], STYLE_ROW_FAIL_NUM if fstat["failed"] > 0 else STYLE_ROW_EVEN_NUM),
            (fstat["status"], STYLE_STATUS_FAIL if is_err else STYLE_STATUS_OK),
            (fstat["period_jst"], STYLE_ROW_EVEN),
        ], height=20)

    return ws


def build_reference_sheet(sst: SharedStringsTable) -> WorksheetBuilder:
    col_widths = [10, 32, 30, 22, 58, 46]
    num_cols = len(col_widths)
    ws = WorksheetBuilder(
        name="項目定義・エラーコード解説",
        sst=sst,
        col_widths=col_widths,
        freeze_row=2,
    )

    ws.add_row(
        [("AWS ALB 接続ログ (Connection Log) 出力項目定義・AWSエラーコード解説リファレンス", STYLE_TITLE)]
        + [("", STYLE_TITLE) for _ in range(num_cols - 1)],
        height=28,
    )
    ws.merge_cells(1, 1, 1, num_cols)

    ws.add_row(
        [("本ブックの「接続ログ一覧」「エラー・失敗ログ一覧」シートに出力される各ヘッダ項目の意味と、AWS公式の接続エラーコード一覧です。", STYLE_SUBTITLE)]
        + [("", STYLE_SUBTITLE) for _ in range(num_cols - 1)],
        height=22,
    )
    ws.merge_cells(2, 1, 2, num_cols)

    ws.add_row([("", STYLE_DEFAULT)], height=10)

    # セクション1: 出力項目一覧
    r_sec1 = ws.add_row([("1. 接続ログ一覧シート 各出力項目（ヘッダ）の内容定義", STYLE_SEC_HEADER)] + [("", STYLE_SEC_HEADER) for _ in range(num_cols - 1)], height=24)
    ws.merge_cells(r_sec1, 1, r_sec1, num_cols)

    ws.add_row([
        ("列番号", STYLE_TBL_HEADER),
        ("Excelヘッダ項目名", STYLE_TBL_HEADER),
        ("AWS公式フィールド名 (位置)", STYLE_TBL_HEADER),
        ("データ型", STYLE_TBL_HEADER),
        ("出力項目の内容・詳細説明", STYLE_TBL_HEADER),
        ("「-」が出力される条件・備考", STYLE_TBL_HEADER),
    ], height=28)

    for idx, col in enumerate(DETAIL_COLUMNS, start=1):
        col_let = col_letter(idx)
        header_first_line = col[2].split("\n")[0]
        s_row = STYLE_ROW_ODD if idx % 2 == 1 else STYLE_ROW_EVEN
        s_center = STYLE_ROW_ODD_CENTER if idx % 2 == 1 else STYLE_ROW_EVEN_CENTER
        ws.add_row([
            (f"{col_let}列 (#{idx})", s_center),
            (header_first_line, s_row),
            (col[4], s_row),
            (col[5], s_center),
            (col[6], s_row),
            (col[7], s_row),
        ], height=22)

    ws.add_row([("", STYLE_DEFAULT)], height=14)

    # セクション2: AWS公式エラーコード一覧
    r_sec2 = ws.add_row([("2. AWS ALB 接続ログ エラー理由コード (Failed:<error_code>) 解説一覧", STYLE_SEC_HEADER)] + [("", STYLE_SEC_HEADER) for _ in range(num_cols - 1)], height=24)
    ws.merge_cells(r_sec2, 1, r_sec2, num_cols)

    ws.add_row([
        ("No.", STYLE_TBL_HEADER),
        ("tls_verify_status 出力値", STYLE_TBL_HEADER),
        ("エラー理由コード (Code)", STYLE_TBL_HEADER),
        ("日本語概要", STYLE_TBL_HEADER),
        ("詳細な原因と確認・対処ポイント", STYLE_TBL_HEADER),
        ("関連する主な確認項目", STYLE_TBL_HEADER),
    ], height=28)

    for idx, (err_code, (short_ja, detail_ja)) in enumerate(AWS_ERROR_DESCRIPTIONS.items(), start=1):
        s_row = STYLE_ROW_ODD if idx % 2 == 1 else STYLE_ROW_EVEN
        s_center = STYLE_ROW_ODD_CENTER if idx % 2 == 1 else STYLE_ROW_EVEN_CENTER
        if "Expired" in err_code or "NotYetValid" in err_code:
            check_col = "leaf_client_cert_validity / cert_validity_jst"
        elif "Crl" in err_code or "Untrusted" in err_code or "Chain" in err_code:
            check_col = "leaf_client_cert_subject / ALBトラストストア・CRL設定"
        elif "Incompatible" in err_code or "Unsupported" in err_code:
            check_col = "tls_protocol / tls_cipher / リスナーセキュリティポリシー"
        else:
            check_col = "client_ip / leaf_client_cert_subject / conn_trace_id"

        ws.add_row([
            (idx, s_center),
            (f"Failed:{err_code}", STYLE_STATUS_FAIL),
            (err_code, s_row),
            (short_ja, s_row),
            (detail_ja, s_row),
            (check_col, s_row),
        ], height=22)

    return ws


def write_xlsx_workbook(output_path: str, worksheets: list, sst: SharedStringsTable):
    """
    OpenXML (ECMA-376) 準拠の .xlsx ファイルを生成する
    """
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # 1. [Content_Types].xml
        ct_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '  <Default Extension="xml" ContentType="application/xml"/>',
            '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
            '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
        ]
        for i in range(1, len(worksheets) + 1):
            ct_parts.append(
                f'  <Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )
        ct_parts.append("</Types>")
        zf.writestr("[Content_Types].xml", "\n".join(ct_parts))

        # 2. _rels/.rels
        rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
        zf.writestr("_rels/.rels", rels_xml)

        # 3. xl/workbook.xml & xl/_rels/workbook.xml.rels
        wb_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '  <sheets>',
        ]
        wb_rels_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        ]
        for i, ws in enumerate(worksheets, start=1):
            safe_name = clean_xml_text(ws.name[:31])
            wb_parts.append(f'    <sheet name="{safe_name}" sheetId="{i}" r:id="rId{i}"/>')
            wb_rels_parts.append(
                f'  <Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
            )
        wb_parts.append("  </sheets>")
        wb_parts.append("</workbook>")

        style_rid = len(worksheets) + 1
        sst_rid = len(worksheets) + 2
        wb_rels_parts.append(
            f'  <Relationship Id="rId{style_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        )
        wb_rels_parts.append(
            f'  <Relationship Id="rId{sst_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        )
        wb_rels_parts.append("</Relationships>")

        zf.writestr("xl/workbook.xml", "\n".join(wb_parts))
        zf.writestr("xl/_rels/workbook.xml.rels", "\n".join(wb_rels_parts))

        # 4. xl/styles.xml
        zf.writestr("xl/styles.xml", build_styles_xml())

        # 5. xl/worksheets/sheet{i}.xml
        for i, ws in enumerate(worksheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", ws.to_xml())

        # 6. xl/sharedStrings.xml
        zf.writestr("xl/sharedStrings.xml", sst.to_xml())


# ==============================================================================
# 4. メイン処理
# ==============================================================================
def main():
    input_dir = Path(sys.argv[1]).resolve()
    output_path = Path(sys.argv[2]).resolve()
    sort_by_time = sys.argv[3] == "1"

    gz_files = []
    for root, _, files in os.walk(input_dir, followlinks=True):
        for fname in files:
            if fname.lower().endswith(".gz"):
                gz_files.append(Path(root) / fname)
    gz_files.sort()
    if not gz_files:
        print(f"[ERROR] .gz ファイルが見つかりません: {input_dir}", file=sys.stderr)
        sys.exit(1)

    all_records = []
    malformed_rows = []
    file_stats = []

    for gz_path in gz_files:
        rel_path = str(gz_path.relative_to(input_dir)).replace("\\", "/")
        size_bytes = gz_path.stat().st_size
        default_elb, default_node_ip = extract_metadata_from_filename(gz_path.name)

        f_records = []
        f_status = "OK"
        try:
            with gzip.open(gz_path, "rt", encoding="utf-8", errors="replace") as f:
                for line_no, raw_line in enumerate(f, start=1):
                    parsed = parse_single_line(raw_line, rel_path, line_no, default_elb, default_node_ip)
                    if parsed is None:
                        continue
                    if parsed.get("_malformed"):
                        malformed_rows.append(parsed)
                    else:
                        f_records.append(parsed)
                        all_records.append(parsed)
        except Exception as e:
            f_status = f"読込エラー ({type(e).__name__}: {e})"
            print(f"[WARN] ファイル読込中にエラーが発生しました: {rel_path} - {e}", file=sys.stderr)

        f_dts = [r["_dt"] for r in f_records if r["_dt"] is not None]
        if f_dts:
            f_period = f"{format_jst(min(f_dts))} ～ {format_jst(max(f_dts))}"
        else:
            f_period = "-"

        file_stats.append({
            "rel_path": rel_path,
            "size_bytes": size_bytes,
            "records": len(f_records),
            "success": sum(1 for r in f_records if r["_category"] == "SUCCESS"),
            "failed": sum(1 for r in f_records if r["_category"] == "FAILED"),
            "status": f_status if f_records or f_status != "OK" else "空ファイル (0件)",
            "period_jst": f_period,
        })

    if sort_by_time:
        all_records.sort(key=lambda r: (r["_sort_key"], r["source_file"]))

    total_records = len(all_records)
    success_records = sum(1 for r in all_records if r["_category"] == "SUCCESS")
    failed_records_list = [r for r in all_records if r["_category"] == "FAILED"]
    failed_records = len(failed_records_list)
    non_tls_records = sum(1 for r in all_records if r["_category"] == "NON_TLS")

    sst = SharedStringsTable()
    worksheets = []

    # シート1: 接続ログ一覧 (1行ずつ全項目をわかりやすく確認できるメインシート)
    # Excelの最大行数(1,048,576行)を考慮し、90万行を超える場合は自動分割
    MAX_ROWS_PER_SHEET = 900000
    if total_records <= MAX_ROWS_PER_SHEET:
        subtitle = (
            f"対象ディレクトリ: {input_dir.name} | 読込gzファイル数: {len(gz_files)}件 | "
            f"総レコード数: {total_records:,}件 (正常[Success]: {success_records:,}件 / "
            f"失敗[Failed]: {failed_records:,}件 / 非TLS等[-]: {non_tls_records:,}件)"
        )
        ws_detail = build_detail_sheet(
            sheet_name="接続ログ一覧",
            title_text="AWS ALB 接続ログ (Connection Log) 1行別詳細一覧レポート",
            subtitle_text=subtitle,
            records=all_records,
            sst=sst,
            empty_message="有効な接続ログレコードが存在しません。",
        )
        worksheets.append(ws_detail)
    else:
        for chunk_idx, start_i in enumerate(range(0, total_records, MAX_ROWS_PER_SHEET), start=1):
            chunk = all_records[start_i : start_i + MAX_ROWS_PER_SHEET]
            sname = "接続ログ一覧" if chunk_idx == 1 else f"接続ログ一覧_{chunk_idx}"
            subtitle = (
                f"分割シート #{chunk_idx} ({start_i + 1:,}～{start_i + len(chunk):,}件目 / 全{total_records:,}件) | "
                f"読込gzファイル数: {len(gz_files)}件"
            )
            ws_detail = build_detail_sheet(
                sheet_name=sname,
                title_text=f"AWS ALB 接続ログ (Connection Log) 1行別詳細一覧レポート ({chunk_idx})",
                subtitle_text=subtitle,
                records=chunk,
                sst=sst,
            )
            worksheets.append(ws_detail)

    # シート2: 分析サマリーレポート
    ws_summary = build_summary_sheet(
        input_dir=str(input_dir),
        file_stats=file_stats,
        records=all_records,
        malformed_rows=malformed_rows,
        sst=sst,
    )
    worksheets.append(ws_summary)

    # シート3: エラー・失敗ログ一覧 (Failed:* のみ抽出)
    fail_subtitle = (
        f"抽出条件: tls_verify_status が Failed:* の接続失敗レコード | "
        f"該当件数: {failed_records:,}件 / 全{total_records:,}件 "
        f"({(failed_records / total_records * 100) if total_records > 0 else 0.0:.2f}%)"
    )
    ws_failed = build_detail_sheet(
        sheet_name="エラー・失敗ログ一覧",
        title_text="AWS ALB 接続ログ エラー・失敗レコード (Failed) 抽出一覧",
        subtitle_text=fail_subtitle,
        records=failed_records_list[:MAX_ROWS_PER_SHEET],
        sst=sst,
        empty_message="エラー・失敗ログ (Failed) はありません（全件正常に処理されています）。",
    )
    worksheets.append(ws_failed)

    # シート4: 項目定義・エラーコード解説
    ws_ref = build_reference_sheet(sst=sst)
    worksheets.append(ws_ref)

    write_xlsx_workbook(str(output_path), worksheets, sst)

    print("--------------------------------------------------------------------------------")
    print("  [解析・集計結果サマリー]")
    print(f"  ・読込 .gz ファイル数    : {len(gz_files):,} ファイル")
    print(f"  ・総接続ログ件数         : {total_records:,} 件")
    print(f"  ・正常接続 (Success)     : {success_records:,} 件 ({(success_records/total_records*100) if total_records else 0:.2f}%)")
    print(f"  ・失敗接続 (Failed:*)    : {failed_records:,} 件 ({(failed_records/total_records*100) if total_records else 0:.2f}%)")
    print(f"  ・非TLS/未設定 (-)       : {non_tls_records:,} 件")
    if malformed_rows:
        print(f"  ・形式不正スキップ行数   : {len(malformed_rows):,} 行")
    print(f"  ・出力シート構成         : {', '.join(ws.name for ws in worksheets)}")
    print(f"  ・出力Excelファイルパス  : {output_path}")
    print("--------------------------------------------------------------------------------")


if __name__ == "__main__":
    main()
PYEOF

log_info "Excelレポートの出力が完了しました: ${OUTPUT_PATH}"
