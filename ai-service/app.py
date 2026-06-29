"""
============================================
PerfumeShop V18 AI Microservice
Flask REST API for AI-powered features:
  - Product Recommendations (collaborative + content-based)
  - Fragrance Matching (quiz-based formula recommendation)
  - Sentiment Analysis (review sentiment)
  - Chatbot Handler (FAQ + AI fallback)
============================================
"""
import os
import logging
from flask import Flask, request, jsonify
from flask_cors import CORS

from recommendation_engine import RecommendationEngine
from fragrance_matcher import FragranceMatcher
from sentiment_analyzer import SentimentAnalyzer
from chatbot_handler import ChatbotHandler

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s'
)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)
CORS(app)

# Initialize service modules
rec_engine = RecommendationEngine()
frag_matcher = FragranceMatcher()
sent_analyzer = SentimentAnalyzer()
chatbot = ChatbotHandler()

# ============================================
# Health Check
# ============================================
@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({
        'status': 'ok',
        'service': 'PerfumeShop AI Microservice',
        'version': 'V18.0',
        'modules': {
            'recommendation': rec_engine.is_ready(),
            'fragrance_matcher': frag_matcher.is_ready(),
            'sentiment': sent_analyzer.is_ready(),
            'chatbot': chatbot.is_ready()
        }
    })

# ============================================
# Recommendation API
# ============================================
@app.route('/api/recommend/personalized', methods=['POST'])
def recommend_personalized():
    """Get personalized product recommendations for a user."""
    try:
        data = request.get_json() or {}
        user_id = data.get('user_id')
        limit = data.get('limit', 10)
        exclude_ids = data.get('exclude_ids', [])
        
        if user_id is None:
            return jsonify({'error': 'Missing user_id'}), 400
        
        recommendations = rec_engine.get_personalized(user_id, limit, exclude_ids)
        return jsonify({'code': 0, 'data': recommendations})
    except Exception as e:
        logger.error(f"Recommendation error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/recommend/similar', methods=['POST'])
def recommend_similar():
    """Get similar products based on a reference product."""
    try:
        data = request.get_json() or {}
        product_id = data.get('product_id')
        limit = data.get('limit', 6)
        
        if product_id is None:
            return jsonify({'error': 'Missing product_id'}), 400
        
        similar = rec_engine.get_similar_products(product_id, limit)
        return jsonify({'code': 0, 'data': similar})
    except Exception as e:
        logger.error(f"Similar recommendation error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/recommend/trending', methods=['GET'])
def recommend_trending():
    """Get trending products."""
    try:
        limit = request.args.get('limit', 8, type=int)
        trending = rec_engine.get_trending(limit)
        return jsonify({'code': 0, 'data': trending})
    except Exception as e:
        logger.error(f"Trending error: {e}")
        return jsonify({'error': str(e)}), 500

# ============================================
# Fragrance Matching API
# ============================================
@app.route('/api/fragrance/match', methods=['POST'])
def fragrance_match():
    """Match fragrances based on user preferences quiz."""
    try:
        data = request.get_json() or {}
        answers = data.get('answers', {})
        
        if not answers:
            return jsonify({'error': 'Missing quiz answers'}), 400
        
        result = frag_matcher.match(answers)
        return jsonify({'code': 0, 'data': result})
    except Exception as e:
        logger.error(f"Fragrance match error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/fragrance/quiz', methods=['GET'])
def fragrance_quiz():
    """Get fragrance quiz questions."""
    questions = frag_matcher.get_quiz_questions()
    return jsonify({'code': 0, 'data': questions})

# ============================================
# Sentiment Analysis API
# ============================================
@app.route('/api/sentiment/analyze', methods=['POST'])
def sentiment_analyze():
    """Analyze sentiment of review text."""
    try:
        data = request.get_json() or {}
        text = data.get('text', '')
        
        if not text.strip():
            return jsonify({'error': 'Missing text'}), 400
        
        result = sent_analyzer.analyze(text)
        return jsonify({'code': 0, 'data': result})
    except Exception as e:
        logger.error(f"Sentiment error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/sentiment/batch', methods=['POST'])
def sentiment_batch():
    """Batch analyze sentiments of multiple reviews."""
    try:
        data = request.get_json() or {}
        texts = data.get('texts', [])
        
        if not texts:
            return jsonify({'error': 'Missing texts array'}), 400
        
        results = [sent_analyzer.analyze(t) for t in texts]
        return jsonify({'code': 0, 'data': results})
    except Exception as e:
        logger.error(f"Batch sentiment error: {e}")
        return jsonify({'error': str(e)}), 500

# ============================================
# Chatbot API
# ============================================
@app.route('/api/chatbot/message', methods=['POST'])
def chatbot_message():
    """Handle chatbot conversation."""
    try:
        data = request.get_json() or {}
        message = data.get('message', '')
        session_id = data.get('session_id', '')
        context = data.get('context', {})
        
        if not message.strip():
            return jsonify({'error': 'Missing message'}), 400
        
        response = chatbot.process_message(message, session_id, context)
        return jsonify({'code': 0, 'data': response})
    except Exception as e:
        logger.error(f"Chatbot error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/chatbot/faq', methods=['GET'])
def chatbot_faq():
    """Get FAQ categories and questions."""
    faqs = chatbot.get_faq_list()
    return jsonify({'code': 0, 'data': faqs})

# ============================================
# Main Entry Point
# ============================================
if __name__ == '__main__':
    port = int(os.environ.get('AI_SERVICE_PORT', 5000))
    debug = os.environ.get('AI_SERVICE_DEBUG', '0') == '1'
    logger.info(f"Starting AI Service on port {port} (debug={debug})")
    app.run(host='0.0.0.0', port=port, debug=debug)
