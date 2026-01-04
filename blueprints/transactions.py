from flask import Blueprint, render_template, request, redirect, url_for, flash
from db import SessionLocal
from models import ImportedTransaction, Account, AccountItem, JournalEntry, GeneralLedger
from datetime import datetime
import csv
import io
from werkzeug.utils import secure_filename
import openpyxl

# Blueprintの作成
transactions_bp = Blueprint('transactions', __name__, url_prefix='/transactions')

# ログイン必須デコレータ（app.pyからインポート）
def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated_function(*args, **kwargs):
        from flask import session, redirect, url_for
        if 'organization_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# 現在の事業所を取得する関数（app.pyからインポート）
def get_current_organization():
    from flask import session
    from models import Organization
    db = SessionLocal()
    try:
        if 'organization_id' in session:
            return db.query(Organization).filter(Organization.id == session['organization_id']).first()
        return None
    finally:
        db.close()

# ========== 取引明細インポート ==========
@transactions_bp.route('/import', methods=['GET'])
@login_required
def transaction_import():
    """取引明細インポートページ"""
    db = SessionLocal()
    try:
        # 登録済み口座を取得
        accounts = db.query(Account).filter(
            Account.is_visible_in_list == 1
        ).order_by(Account.id.asc()).all()
        
        return render_template(
            'transactions/import.html',
            accounts=accounts
        )
    finally:
        db.close()

@transactions_bp.route('/import', methods=['POST'])
@login_required
def transaction_import_upload():
    """取引明細のアップロード処理"""
    db = SessionLocal()
    try:
        # フォームデータの取得
        account_id = request.form.get('account_id', type=int)
        file = request.files.get('file')
        
        if not account_id or not file:
            flash('口座とファイルを選択してください', 'error')
            return redirect(url_for('transactions.transaction_import'))
        
        # 口座情報を取得
        account = db.query(Account).filter(Account.id == account_id).first()
        if not account:
            flash('選択された口座が見つかりません', 'error')
            return redirect(url_for('transactions.transaction_import'))
        
        # ファイルの拡張子を確認
        filename = secure_filename(file.filename)
        file_ext = filename.rsplit('.', 1)[1].lower() if '.' in filename else ''
        
        # 現在の事業所を取得
        current_org = get_current_organization()
        if not current_org:
            flash('事業所が見つかりません', 'error')
            return redirect(url_for('transactions.transaction_import'))
        
        imported_count = 0
        
        # CSVファイルの処理
        if file_ext == 'csv':
            # CSVファイルを読み込み
            stream = io.StringIO(file.stream.read().decode('utf-8-sig'), newline=None)
            csv_reader = csv.DictReader(stream)
            
            for row in csv_reader:
                # 取引日のパース
                transaction_date = row.get('取引日', '').strip()
                if not transaction_date:
                    continue
                
                # 日付フォーマットの変換（YYYY-MM-DD形式に統一）
                try:
                    date_obj = datetime.strptime(transaction_date, '%Y-%m-%d')
                    transaction_date = date_obj.strftime('%Y-%m-%d')
                except:
                    try:
                        date_obj = datetime.strptime(transaction_date, '%Y/%m/%d')
                        transaction_date = date_obj.strftime('%Y-%m-%d')
                    except:
                        continue
                
                # 摘要
                description = row.get('摘要', '').strip()
                
                # 入金金額
                income_str = row.get('入金金額', '0').strip().replace(',', '')
                income_amount = int(income_str) if income_str and income_str != '' else 0
                
                # 出金金額
                expense_str = row.get('出金金額', '0').strip().replace(',', '')
                expense_amount = int(expense_str) if expense_str and expense_str != '' else 0
                
                # ImportedTransactionを作成
                imported_transaction = ImportedTransaction(
                    organization_id=current_org.id,
                    account_name=account.account_name,
                    transaction_date=transaction_date,
                    description=description,
                    income_amount=income_amount,
                    expense_amount=expense_amount,
                    status=0,  # 未処理
                    imported_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                )
                db.add(imported_transaction)
                imported_count += 1
        
        # Excelファイルの処理
        elif file_ext in ['xlsx', 'xls']:
            # Excelファイルを読み込み
            workbook = openpyxl.load_workbook(file)
            sheet = workbook.active
            
            # ヘッダー行を取得（1行目）
            headers = [cell.value for cell in sheet[1]]
            
            # データ行を処理（2行目以降）
            for row in sheet.iter_rows(min_row=2, values_only=True):
                row_dict = dict(zip(headers, row))
                
                # 取引日のパース
                transaction_date = row_dict.get('取引日', '')
                if not transaction_date:
                    continue
                
                # 日付型の場合はstrftimeで変換
                if isinstance(transaction_date, datetime):
                    transaction_date = transaction_date.strftime('%Y-%m-%d')
                else:
                    transaction_date = str(transaction_date).strip()
                    try:
                        date_obj = datetime.strptime(transaction_date, '%Y-%m-%d')
                        transaction_date = date_obj.strftime('%Y-%m-%d')
                    except:
                        try:
                            date_obj = datetime.strptime(transaction_date, '%Y/%m/%d')
                            transaction_date = date_obj.strftime('%Y-%m-%d')
                        except:
                            continue
                
                # 摘要
                description = str(row_dict.get('摘要', '')).strip()
                
                # 入金金額
                income_amount = int(row_dict.get('入金金額', 0) or 0)
                
                # 出金金額
                expense_amount = int(row_dict.get('出金金額', 0) or 0)
                
                # ImportedTransactionを作成
                imported_transaction = ImportedTransaction(
                    organization_id=current_org.id,
                    account_name=account.account_name,
                    transaction_date=transaction_date,
                    description=description,
                    income_amount=income_amount,
                    expense_amount=expense_amount,
                    status=0,  # 未処理
                    imported_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                )
                db.add(imported_transaction)
                imported_count += 1
        
        else:
            flash('CSVまたはExcelファイルを選択してください', 'error')
            return redirect(url_for('transactions.transaction_import'))
        
        # データベースにコミット
        db.commit()
        
        flash(f'{imported_count}件の取引明細をインポートしました', 'success')
        return redirect(url_for('transactions.imported_transactions_list'))
    
    except Exception as e:
        db.rollback()
        flash(f'インポート中にエラーが発生しました: {str(e)}', 'error')
        return redirect(url_for('transactions.transaction_import'))
    finally:
        db.close()

# ========== 取引明細一覧 ==========
@transactions_bp.route('/imported', methods=['GET'])
@login_required
def imported_transactions_list():
    """インポートされた取引明細の一覧"""
    db = SessionLocal()
    try:
        # 検索条件
        account_name = request.args.get('account_name', '', type=str)
        account_item_id = request.args.get('account_item_id', type=int)  # 動定科目IDフィルター
        status = request.args.get('status', '', type=str)
        date_from = request.args.get('date_from', '', type=str)
        date_to = request.args.get('date_to', '', type=str)
        
        # 現在の事業所を取得
        current_org = get_current_organization()
        if not current_org:
            flash('事業所が見つかりません', 'error')
            return redirect(url_for('index'))
        
        # クエリ構築
        query = db.query(ImportedTransaction).filter(
            ImportedTransaction.organization_id == current_org.id
        )
        
        # 口座名でフィルタ
        if account_name:
            query = query.filter(ImportedTransaction.account_name == account_name)
        
        # 動定科目IDでフィルタ（account_item_idが指定された場合、その動定科目に対応する口座を検索）
        if account_item_id:
            # account_item_idに対応する口座を取得
            account = db.query(Account).filter(Account.account_item_id == account_item_id).first()
            if account:
                query = query.filter(ImportedTransaction.account_name == account.account_name)
        
        # ステータスでフィルタ
        if status:
            query = query.filter(ImportedTransaction.status == int(status))
        
        # 日付範囲でフィルタ
        if date_from:
            query = query.filter(ImportedTransaction.transaction_date >= date_from)
        if date_to:
            query = query.filter(ImportedTransaction.transaction_date <= date_to)
        
        # 取引日降順でソート
        transactions = query.order_by(ImportedTransaction.transaction_date.desc()).all()
        
        # 口座一覧を取得（フィルタ用）
        accounts = db.query(Account).filter(
            Account.is_visible_in_list == 1
        ).order_by(Account.id.asc()).all()
        
        return render_template(
            'transactions/list.html',
            transactions=transactions,
            accounts=accounts,
            account_name=account_name,
            account_item_id=account_item_id,
            status=status,
            date_from=date_from,
            date_to=date_to
        )
    finally:
        db.close()

# ========== 取引明細編集 ==========
@transactions_bp.route('/imported/<int:id>/edit', methods=['GET'])
@login_required
def imported_transaction_edit(id):
    """取引明細の編集ページ"""
    db = SessionLocal()
    try:
        # 取引明細を取得
        transaction = db.query(ImportedTransaction).filter(ImportedTransaction.id == id).first()
        if not transaction:
            flash('取引明細が見つかりません', 'error')
            return redirect(url_for('transactions.imported_transactions_list'))
        
        # 勘定科目一覧を取得
        account_items = db.query(AccountItem).order_by(AccountItem.id.asc()).all()
        
        return render_template(
            'transactions/edit.html',
            transaction=transaction,
            account_items=account_items,
            selected_account_item_id=transaction.account_item_id if transaction.account_item_id else None
        )
    finally:
        db.close()

@transactions_bp.route('/imported/<int:id>/edit', methods=['POST'])
@login_required
def imported_transaction_update(id):
    """取引明細の更新処理"""
    db = SessionLocal()
    try:
        # 取引明細を取得
        transaction = db.query(ImportedTransaction).filter(ImportedTransaction.id == id).first()
        if not transaction:
            flash('取引明細が見つかりません', 'error')
            return redirect(url_for('transactions.imported_transactions_list'))
        
        # フォームデータの取得
        account_item_id = request.form.get('account_item_id', type=int)
        description = request.form.get('description', '').strip()
        counterparty_id = request.form.get('counterparty_id', type=int) or None
        department_id = request.form.get('department_id', type=int) or None
        item_id = request.form.get('item_id', type=int) or None
        project_tag_id = request.form.get('project_tag_id', type=int) or None
        memo_tag_id = request.form.get('memo_tag_id', type=int) or None
        
        if not account_item_id:
            flash('勘定科目を選択してください', 'error')
            return redirect(url_for('transactions.imported_transaction_edit', id=id))
        
        # 勘定科目を取得
        account_item = db.query(AccountItem).filter(AccountItem.id == account_item_id).first()
        if not account_item:
            flash('選択された勘定科目が見つかりません', 'error')
            return redirect(url_for('transactions.imported_transaction_edit', id=id))
        
        # 口座の勘定科目IDを取得（普通預金など）
        account = db.query(Account).filter(Account.account_name == transaction.account_name).first()
        if not account:
            flash('口座情報が見つかりません', 'error')
            return redirect(url_for('transactions.imported_transaction_edit', id=id))
        
        # 口座に対応する勘定科目を取得（例: 普通預金）
        account_account_item = db.query(AccountItem).filter(
            AccountItem.account_name.like(f'%{account.account_name.split()[0]}%')
        ).first()
        
        if not account_account_item:
            # デフォルトで「普通預金」を使用
            account_account_item = db.query(AccountItem).filter(
                AccountItem.account_name == '普通預金'
            ).first()
        
        if not account_account_item:
            flash('口座に対応する勘定科目が見つかりません', 'error')
            return redirect(url_for('transactions.imported_transaction_edit', id=id))
        
        # 現在の事業所を取得
        current_org = get_current_organization()
        
        # JournalEntryを作成
        if transaction.income_amount > 0:
            # 入金の場合
            journal_entry = JournalEntry(
                organization_id=current_org.id,
                transaction_date=transaction.transaction_date,
                debit_account_item_id=account_account_item.id,  # 借方: 口座
                credit_account_item_id=account_item.id,  # 貸方: 選択された勘定科目
                debit_amount=transaction.income_amount,
                credit_amount=transaction.income_amount,
                summary=description or transaction.description,
                created_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                updated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            )
        else:
            # 出金の場合
            journal_entry = JournalEntry(
                organization_id=current_org.id,
                transaction_date=transaction.transaction_date,
                debit_account_item_id=account_item.id,  # 借方: 選択された勘定科目
                credit_account_item_id=account_account_item.id,  # 貸方: 口座
                debit_amount=transaction.expense_amount,
                credit_amount=transaction.expense_amount,
                summary=description or transaction.description,
                created_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                updated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            )
        
        db.add(journal_entry)
        db.flush()  # IDを取得するためにflush
        
        # GeneralLedgerにも登録
        general_ledger_entry = GeneralLedger(
            organization_id=current_org.id,
            transaction_date=transaction.transaction_date,
            debit_account_item_id=journal_entry.debit_account_item_id,
            debit_amount=journal_entry.debit_amount,
            debit_tax_category_id=journal_entry.debit_tax_category_id,
            credit_account_item_id=journal_entry.credit_account_item_id,
            credit_amount=journal_entry.credit_amount,
            credit_tax_category_id=journal_entry.credit_tax_category_id,
            summary=journal_entry.summary,
            remarks=journal_entry.remarks,
            counterparty_id=counterparty_id,
            department_id=department_id,
            item_id=item_id,
            project_tag_id=project_tag_id,
            memo_tag_id=memo_tag_id,
            source_type='imported_transaction',
            source_id=transaction.id,
            created_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            updated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
        db.add(general_ledger_entry)
        
        # ImportedTransactionのステータスを更新
        transaction.status = 1  # 処理済み
        transaction.journal_entry_id = journal_entry.id
        transaction.account_item_id = account_item_id
        if description:
            transaction.description = description
        
        db.commit()
        
        flash('取引明細を登録しました', 'success')
        return redirect(url_for('transactions.imported_transactions_list'))
    
    except Exception as e:
        db.rollback()
        flash(f'登録中にエラーが発生しました: {str(e)}', 'error')
        return redirect(url_for('transactions.imported_transaction_edit', id=id))
    finally:
        db.close()


@transactions_bp.route('/imported/<int:id>/delete', methods=['POST'])
@login_required
def imported_transaction_delete(id):
    """取引明細の削除処理"""
    db = SessionLocal()
    try:
        # 取引明細を取得
        transaction = db.query(ImportedTransaction).filter(ImportedTransaction.id == id).first()
        if not transaction:
            from flask import jsonify
            return jsonify({'success': False, 'message': '取引明細が見つかりません'}), 404
        
        # 関連する仕訳データを削除
        general_ledger_entries = db.query(GeneralLedger).filter(
            GeneralLedger.source_type == 'imported_transaction',
            GeneralLedger.source_id == id
        ).all()
        
        for entry in general_ledger_entries:
            db.delete(entry)
        
        # 取引明細を削除
        db.delete(transaction)
        db.commit()
        
        from flask import jsonify, flash
        flash('取引を解除しました', 'success')
        return jsonify({'success': True, 'message': '取引を解除しました'})
    
    except Exception as e:
        db.rollback()
        from flask import jsonify
        return jsonify({'success': False, 'message': f'削除中にエラーが発生しました: {str(e)}'}), 500
    finally:
        db.close()
