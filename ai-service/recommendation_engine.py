"""
============================================
PerfumeShop V18 - Recommendation Engine
Collaborative filtering + Content-based recommendations
============================================
"""
import logging
import math
from collections import defaultdict
from typing import List, Dict, Optional

logger = logging.getLogger(__name__)


class RecommendationEngine:
    """Product recommendation engine using collaborative & content-based filtering."""
    
    def __init__(self):
        self._ready = True
        # In production, load user-item matrix from database
        self._user_history = defaultdict(set)  # user_id -> set of product_ids
        self._product_features = {}  # product_id -> feature vector
        self._trending_scores = {}  # product_id -> trending score
        
    def is_ready(self) -> bool:
        return self._ready
    
    def update_user_history(self, user_id: int, product_ids: List[int]):
        """Update user's purchase/browse history."""
        self._user_history[user_id].update(product_ids)
    
    def update_product_features(self, product_id: int, features: Dict[str, float]):
        """Update product feature vector for content-based similarity."""
        self._product_features[product_id] = features
    
    def get_personalized(self, user_id: int, limit: int = 10, exclude_ids: List[int] = None) -> List[Dict]:
        """
        Get personalized product recommendations for a user.
        Combines collaborative filtering with content-based recommendations.
        
        Args:
            user_id: User identifier
            limit: Maximum number of recommendations
            exclude_ids: Product IDs to exclude from results
        
        Returns:
            List of dicts with product_id, score, reason
        """
        exclude_set = set(exclude_ids or [])
        user_products = self._user_history.get(user_id, set())
        exclude_set.update(user_products)
        
        scores = defaultdict(float)
        
        # 1. Collaborative: Find similar users and recommend their products
        similar_users = self._find_similar_users(user_id, top_k=20)
        for sim_user, sim_score in similar_users:
            for pid in self._user_history.get(sim_user, set()):
                if pid not in exclude_set:
                    scores[pid] += sim_score * 1.0
        
        # 2. Content-based: Recommend products similar to user's history
        if user_products and self._product_features:
            user_profile = self._compute_user_profile(user_products)
            for pid, features in self._product_features.items():
                if pid not in exclude_set:
                    content_score = self._cosine_similarity(user_profile, features)
                    scores[pid] += content_score * 0.5
        
        # 3. Boost trending items
        for pid, trend_score in self._trending_scores.items():
            if pid not in exclude_set:
                scores[pid] += trend_score * 0.3
        
        # Sort and return top results
        sorted_items = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        results = []
        for pid, score in sorted_items[:limit]:
            results.append({
                'product_id': pid,
                'score': round(float(score), 4),
                'reason': self._get_recommendation_reason(pid, user_products)
            })
        
        return results
    
    def get_similar_products(self, product_id: int, limit: int = 6) -> List[Dict]:
        """Find products similar to the given product."""
        if product_id not in self._product_features:
            return []
        
        target_features = self._product_features[product_id]
        similarities = []
        
        for pid, features in self._product_features.items():
            if pid != product_id:
                sim = self._cosine_similarity(target_features, features)
                similarities.append((pid, sim))
        
        similarities.sort(key=lambda x: x[1], reverse=True)
        return [
            {'product_id': pid, 'similarity': round(float(sim), 4)}
            for pid, sim in similarities[:limit]
        ]
    
    def get_trending(self, limit: int = 8) -> List[Dict]:
        """Get trending products."""
        sorted_trending = sorted(
            self._trending_scores.items(),
            key=lambda x: x[1], reverse=True
        )
        return [
            {'product_id': pid, 'trend_score': round(float(score), 4)}
            for pid, score in sorted_trending[:limit]
        ]
    
    def _find_similar_users(self, user_id: int, top_k: int = 20) -> List[tuple]:
        """Find users with similar purchase history."""
        target_products = self._user_history.get(user_id, set())
        if not target_products:
            return []
        
        similarities = []
        for uid, products in self._user_history.items():
            if uid != user_id and products:
                intersection = len(target_products & products)
                if intersection > 0:
                    union = len(target_products | products)
                    jaccard = intersection / union if union > 0 else 0
                    similarities.append((uid, jaccard))
        
        similarities.sort(key=lambda x: x[1], reverse=True)
        return similarities[:top_k]
    
    def _compute_user_profile(self, product_ids: set) -> Dict[str, float]:
        """Compute user preference profile from product features."""
        profile = defaultdict(float)
        count = 0
        for pid in product_ids:
            if pid in self._product_features:
                for feat, val in self._product_features[pid].items():
                    profile[feat] += val
                count += 1
        
        if count > 0:
            for feat in profile:
                profile[feat] /= count
        
        return dict(profile)
    
    def _cosine_similarity(self, vec1: Dict[str, float], vec2: Dict[str, float]) -> float:
        """Compute cosine similarity between two feature vectors."""
        all_keys = set(vec1.keys()) | set(vec2.keys())
        dot_product = sum(vec1.get(k, 0) * vec2.get(k, 0) for k in all_keys)
        norm1 = math.sqrt(sum(v ** 2 for v in vec1.values()))
        norm2 = math.sqrt(sum(v ** 2 for v in vec2.values()))
        
        if norm1 == 0 or norm2 == 0:
            return 0.0
        return dot_product / (norm1 * norm2)
    
    def _get_recommendation_reason(self, product_id: int, user_products: set) -> str:
        """Generate human-readable recommendation reason."""
        reasons = []
        if product_id in self._trending_scores and self._trending_scores.get(product_id, 0) > 0.5:
            reasons.append("热门趋势")
        if self._product_features:
            reasons.append("与你喜爱的产品相似")
        if user_products:
            reasons.append("其他用户也喜欢")
        return "、".join(reasons) if reasons else "综合推荐"
