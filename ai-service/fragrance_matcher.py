"""
============================================
PerfumeShop V18 - Fragrance Matcher
Quiz-based fragrance formula recommendation
============================================
"""
import logging
from typing import List, Dict, Any

logger = logging.getLogger(__name__)


class FragranceMatcher:
    """Matches users to fragrance formulas based on preference quiz."""
    
    # Fragrance note families with characteristics
    NOTE_FAMILIES = {
        'floral': {'keywords': ['花香', '玫瑰', '茉莉', '百合', '温柔', '浪漫', '女性'],
                   'weight': 0.9},
        'citrus': {'keywords': ['柑橘', '柠檬', '清新', '活力', '阳光', '清爽', '夏天'],
                   'weight': 0.9},
        'woody': {'keywords': ['木质', '檀香', '雪松', '沉稳', '成熟', '中性', '秋天'],
                  'weight': 0.8},
        'oriental': {'keywords': ['东方', '琥珀', '香草', '神秘', '性感', '浓郁', '夜晚'],
                     'weight': 0.8},
        'fresh': {'keywords': ['海洋', '水生', '绿叶', '运动', '干净', '春天', '日常'],
                  'weight': 0.85},
        'fruity': {'keywords': ['果香', '桃子', '莓果', '甜美', '活泼', '年轻', '派对'],
                   'weight': 0.75},
        'green': {'keywords': ['青草', '绿茶', '自然', '素雅', '文艺', '中性'],
                  'weight': 0.7},
    }
    
    # Occasion mapping
    OCCASION_MAP = {
        'daily': ['fresh', 'citrus', 'green'],
        'work': ['woody', 'green', 'fresh'],
        'date': ['floral', 'oriental', 'fruity'],
        'party': ['oriental', 'fruity', 'floral'],
        'sport': ['fresh', 'citrus'],
        'formal': ['woody', 'oriental', 'floral'],
    }
    
    # Season mapping
    SEASON_MAP = {
        'spring': ['floral', 'green', 'fruity'],
        'summer': ['citrus', 'fresh', 'fruity'],
        'autumn': ['woody', 'oriental'],
        'winter': ['oriental', 'woody', 'floral'],
    }
    
    # Gender preference mapping
    GENDER_MAP = {
        'female': ['floral', 'fruity', 'oriental'],
        'male': ['woody', 'fresh', 'citrus'],
        'unisex': ['citrus', 'green', 'woody'],
    }
    
    QUIZ_QUESTIONS = [
        {
            'id': 'style',
            'question': '你更喜欢哪种风格的香水？',
            'options': [
                {'value': 'floral', 'label': '花香调 - 浪漫优雅'},
                {'value': 'fresh', 'label': '清新调 - 干净自然'},
                {'value': 'woody', 'label': '木质调 - 沉稳大气'},
                {'value': 'oriental', 'label': '东方调 - 神秘性感'},
                {'value': 'citrus', 'label': '柑橘调 - 活力阳光'},
            ]
        },
        {
            'id': 'occasion',
            'question': '你主要在什么场合使用香水？',
            'options': [
                {'value': 'daily', 'label': '日常通勤'},
                {'value': 'work', 'label': '工作办公'},
                {'value': 'date', 'label': '约会聚会'},
                {'value': 'party', 'label': '派对晚宴'},
                {'value': 'sport', 'label': '运动休闲'},
            ]
        },
        {
            'id': 'season',
            'question': '你最喜欢哪个季节？',
            'options': [
                {'value': 'spring', 'label': '春天 🌸'},
                {'value': 'summer', 'label': '夏天 ☀️'},
                {'value': 'autumn', 'label': '秋天 🍂'},
                {'value': 'winter', 'label': '冬天 ❄️'},
            ]
        },
        {
            'id': 'gender',
            'question': '你偏好的香水类型？',
            'options': [
                {'value': 'female', 'label': '女士香水'},
                {'value': 'male', 'label': '男士香水'},
                {'value': 'unisex', 'label': '中性香水'},
            ]
        },
        {
            'id': 'intensity',
            'question': '你喜欢多浓的香水？',
            'options': [
                {'value': 'light', 'label': '淡雅清新 (EDT)'},
                {'value': 'medium', 'label': '适中持久 (EDP)'},
                {'value': 'strong', 'label': '浓郁持久 (Parfum)'},
            ]
        },
        {
            'id': 'budget',
            'question': '你的预算范围？',
            'options': [
                {'value': 'entry', 'label': '入门级 (200-500元)'},
                {'value': 'mid', 'label': '中端 (500-1500元)'},
                {'value': 'premium', 'label': '高端 (1500元以上)'},
            ]
        },
    ]
    
    def __init__(self):
        self._ready = True
        
    def is_ready(self) -> bool:
        return self._ready
    
    def get_quiz_questions(self) -> List[Dict]:
        """Return the fragrance quiz questions."""
        return self.QUIZ_QUESTIONS
    
    def match(self, answers: Dict[str, str]) -> Dict[str, Any]:
        """
        Match fragrance preferences based on quiz answers.
        
        Args:
            answers: Dict with keys like 'style', 'occasion', 'season', 'gender', 'intensity', 'budget'
        
        Returns:
            Dict with matched_note_families, recommended_notes, and match_scores
        """
        scores = {}
        
        # Score from direct style preference
        style = answers.get('style', '')
        if style in self.NOTE_FAMILIES:
            self._add_score(scores, style, 2.0)
        
        # Score from occasion
        occasion = answers.get('occasion', '')
        if occasion in self.OCCASION_MAP:
            for note_family in self.OCCASION_MAP[occasion]:
                self._add_score(scores, note_family, 1.5)
        
        # Score from season
        season = answers.get('season', '')
        if season in self.SEASON_MAP:
            for note_family in self.SEASON_MAP[season]:
                self._add_score(scores, note_family, 1.0)
        
        # Score from gender
        gender = answers.get('gender', '')
        if gender in self.GENDER_MAP:
            for note_family in self.GENDER_MAP[gender]:
                self._add_score(scores, note_family, 1.2)
        
        # Apply family weights
        for family in list(scores.keys()):
            if family in self.NOTE_FAMILIES:
                scores[family] *= self.NOTE_FAMILIES[family]['weight']
        
        # Sort by score
        sorted_families = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        
        # Map to specific note recommendations
        note_recommendations = self._get_note_recommendations(sorted_families)
        
        # Intensity mapping
        intensity = answers.get('intensity', 'medium')
        intensity_notes = {
            'light': '建议选择EDT淡香水，清新不张扬',
            'medium': '建议选择EDP淡香精，持久适中',
            'strong': '建议选择Parfum浓香精，持久浓郁',
        }
        
        return {
            'matched_families': [
                {'family': f, 'score': round(s, 2), 
                 'keywords': self.NOTE_FAMILIES.get(f, {}).get('keywords', [])}
                for f, s in sorted_families[:3]
            ],
            'recommended_notes': note_recommendations,
            'intensity_advice': intensity_notes.get(intensity, ''),
            'budget_level': answers.get('budget', 'mid'),
        }
    
    def _add_score(self, scores: Dict[str, float], family: str, value: float):
        """Add score to a note family."""
        if family in scores:
            scores[family] += value
        else:
            scores[family] = value
    
    def _get_note_recommendations(self, sorted_families: List[tuple]) -> Dict[str, List[str]]:
        """Map note families to specific top/middle/base note recommendations."""
        note_map = {
            'floral': {'top': ['佛手柑', '粉红胡椒'], 'middle': ['玫瑰', '茉莉', '鸢尾花'], 'base': ['麝香', '琥珀']},
            'citrus': {'top': ['柠檬', '葡萄柚', '佛手柑'], 'middle': ['橙花', '薄荷'], 'base': ['雪松', '白麝香']},
            'woody': {'top': ['香柠檬', '胡椒'], 'middle': ['雪松', '檀香木'], 'base': ['香根草', '广藿香', '皮革']},
            'oriental': {'top': ['肉桂', '小豆蔻'], 'middle': ['琥珀', '香草', '零陵香豆'], 'base': ['檀香', '麝香', '广藿香']},
            'fresh': {'top': ['柑橘', '海洋'], 'middle': ['薰衣草', '迷迭香'], 'base': ['白麝香', '苔藓']},
            'fruity': {'top': ['桃子', '黑加仑'], 'middle': ['玫瑰', '紫罗兰'], 'base': ['香草', '麝香']},
            'green': {'top': ['佛手柑', '绿叶'], 'middle': ['绿茶', '茉莉'], 'base': ['白麝香', '雪松']},
        }
        
        recommendations = {'top': [], 'middle': [], 'base': []}
        
        # Take top 2-3 families and collect their notes
        for family, score in sorted_families[:2]:
            if family in note_map:
                for layer in ['top', 'middle', 'base']:
                    for note in note_map[family][layer]:
                        if note not in recommendations[layer]:
                            recommendations[layer].append(note)
        
        # Limit to 3 per layer
        for layer in recommendations:
            recommendations[layer] = recommendations[layer][:3]
        
        return recommendations
