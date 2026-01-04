"""認証とマルチテナント対応のヘルパー関数"""
from flask import session

def get_current_organization_id():
    """現在ログイン中の事業所IDを取得"""
    return session.get('organization_id')

def add_organization_filter(query, model):
    """
    クエリに事業所フィルターを追加
    
    Args:
        query: SQLAlchemyのクエリオブジェクト
        model: モデルクラス（organization_id属性を持つ）
    
    Returns:
        フィルター適用後のクエリ
    """
    org_id = get_current_organization_id()
    if org_id and hasattr(model, 'organization_id'):
        return query.filter(model.organization_id == org_id)
    return query

def set_organization_id(obj):
    """
    オブジェクトに現在の事業所IDを設定
    
    Args:
        obj: データベースモデルのインスタンス
    """
    org_id = get_current_organization_id()
    if org_id and hasattr(obj, 'organization_id'):
        obj.organization_id = org_id
