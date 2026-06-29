"""
============================================
PerfumeShop V18 - Sentiment Analyzer
Review sentiment analysis using keyword-based approach
(with jieba for Chinese text segmentation)
============================================
"""
import logging
import re
from typing import Dict, List, Any

logger = logging.getLogger(__name__)

try:
    import jieba
    JIEBA_AVAILABLE = True
except ImportError:
    JIEBA_AVAILABLE = False
    logger.warning("jieba not installed, using basic word matching")


class SentimentAnalyzer:
    """Analyzes sentiment of product review texts."""
    
    # Sentiment word dictionaries (Chinese)
    POSITIVE_WORDS = {
        '喜欢', '不错', '好闻', '香', '推荐', '满意', '棒', '赞', '完美', '惊喜',
        '优雅', '清新', '持久', '高级', '自然', '舒服', '柔和', '温柔', '迷人',
        '好评', '超值', '回购', '性价比', '经典', '独特', '精致', '好', '爱',
        'nice', 'love', 'good', 'great', 'beautiful', 'excellent', 'perfect',
        'wonderful', 'amazing', 'fantastic', 'best', 'lovely'
    }
    
    NEGATIVE_WORDS = {
        '失望', '不好', '难闻', '刺鼻', '不持久', '太淡', '太浓', '不值', '差',
        '后悔', '过敏', '头晕', '恶心', '酒精', '假货', '劣质', '粗糙', '呛',
        'bad', 'poor', 'terrible', 'awful', 'worst', 'horrible', 'dislike',
        'cheap', 'fake', 'weak', 'strong', 'headache', 'allergy'
    }
    
    INTENSIFIERS = {
        '非常': 2.0, '特别': 2.0, '超级': 2.5, '极其': 2.5, '太': 1.8,
        '很': 1.5, '挺': 1.3, '有点': 0.7, '稍微': 0.6, '略微': 0.6,
        'very': 2.0, 'extremely': 2.5, 'really': 1.8, 'quite': 1.5,
        'so': 1.5, 'absolutely': 2.5
    }
    
    NEGATORS = {'不', '没', '无', '非', 'not', 'no', "don't", "doesn't", "didn't"}
    
    def __init__(self):
        self._ready = True
        
    def is_ready(self) -> bool:
        return self._ready
    
    def analyze(self, text: str) -> Dict[str, Any]:
        """
        Analyze sentiment of a single text.
        
        Returns:
            Dict with sentiment ('positive'/'negative'/'neutral'),
            score (-1.0 to 1.0), keywords, and confidence.
        """
        text = text.strip()
        if not text:
            return {
                'sentiment': 'neutral',
                'score': 0.0,
                'confidence': 1.0,
                'keywords': [],
                'summary': ''
            }
        
        # Tokenize
        words = self._tokenize(text)
        
        # Count sentiments
        pos_count = 0
        neg_count = 0
        pos_keywords = []
        neg_keywords = []
        
        for i, word in enumerate(words):
            # Check for negator before sentiment word
            negator_before = False
            if i > 0 and words[i-1] in self.NEGATORS:
                negator_before = True
            
            # Check for intensifier
            intensifier = 1.0
            if i > 0 and words[i-1] in self.INTENSIFIERS:
                intensifier = self.INTENSIFIERS[words[i-1]]
            
            if word in self.POSITIVE_WORDS:
                if negator_before:
                    neg_count += intensifier
                    neg_keywords.append(f"不{word}")
                else:
                    pos_count += intensifier
                    pos_keywords.append(word)
            elif word in self.NEGATIVE_WORDS:
                if negator_before:
                    pos_count += intensifier * 0.7
                    pos_keywords.append(f"不{word}")
                else:
                    neg_count += intensifier
                    neg_keywords.append(word)
        
        # Calculate score
        total = pos_count + neg_count
        if total == 0:
            score = 0.0
            sentiment = 'neutral'
        else:
            score = (pos_count - neg_count) / (pos_count + neg_count)
        
        # Determine sentiment
        if score > 0.2:
            sentiment = 'positive'
        elif score < -0.2:
            sentiment = 'negative'
        else:
            sentiment = 'neutral'
        
        # Confidence based on keyword count
        keyword_count = len(pos_keywords) + len(neg_keywords)
        confidence = min(1.0, keyword_count / 5.0) if keyword_count > 0 else 0.3
        
        # Generate summary
        summary = self._generate_summary(sentiment, pos_keywords[:5], neg_keywords[:5])
        
        return {
            'sentiment': sentiment,
            'score': round(score, 3),
            'confidence': round(confidence, 3),
            'keywords': {
                'positive': pos_keywords[:5],
                'negative': neg_keywords[:5]
            },
            'summary': summary
        }
    
    def _tokenize(self, text: str) -> List[str]:
        """Tokenize text into words/phrases."""
        # Remove special characters but keep Chinese
        text = re.sub(r'[^\u4e00-\u9fff\w\s]', ' ', text)
        
        if JIEBA_AVAILABLE and re.search(r'[\u4e00-\u9fff]', text):
            # Use jieba for Chinese text
            return [w.strip() for w in jieba.cut(text) if w.strip()]
        else:
            # Simple whitespace split for English
            return text.lower().split()
    
    def _generate_summary(self, sentiment: str, pos_words: List[str], neg_words: List[str]) -> str:
        """Generate a human-readable summary."""
        if sentiment == 'positive':
            if pos_words:
                return f"用户评价积极，关键词: {', '.join(pos_words[:3])}"
            return "用户评价偏正面"
        elif sentiment == 'negative':
            if neg_words:
                return f"用户评价消极，关键词: {', '.join(neg_words[:3])}"
            return "用户评价偏负面"
        else:
            return "用户评价中性"
