"""
============================================
PerfumeShop V18 - Chatbot Handler
FAQ knowledge base + rule engine for customer service
============================================
"""
import logging
import re
from typing import Dict, List, Any, Optional
from collections import defaultdict

logger = logging.getLogger(__name__)


class ChatbotHandler:
    """Rule-based chatbot with FAQ knowledge base."""
    
    # Session memory for conversation context
    _sessions = defaultdict(list)
    MAX_SESSION_HISTORY = 20
    
    def __init__(self):
        self._ready = True
        self._faq = self._build_faq()
        
    def is_ready(self) -> bool:
        return self._ready
    
    def get_faq_list(self) -> List[Dict]:
        """Return FAQ categories and questions."""
        return [
            {
                'category': 'order',
                'name': '订单相关',
                'questions': [q['question'] for q in self._faq.get('order', [])]
            },
            {
                'category': 'shipping',
                'name': '配送相关',
                'questions': [q['question'] for q in self._faq.get('shipping', [])]
            },
            {
                'category': 'product',
                'name': '产品相关',
                'questions': [q['question'] for q in self._faq.get('product', [])]
            },
            {
                'category': 'customize',
                'name': '定制流程',
                'questions': [q['question'] for q in self._faq.get('customize', [])]
            },
            {
                'category': 'account',
                'name': '账户相关',
                'questions': [q['question'] for q in self._faq.get('account', [])]
            },
        ]
    
    def process_message(self, message: str, session_id: str = '', context: Dict = None) -> Dict[str, Any]:
        """
        Process a user message and return a response.
        
        Returns:
            Dict with reply, category, needs_human, suggestions
        """
        message = message.strip()
        context = context or {}
        
        # Store in session history
        if session_id:
            self._sessions[session_id].append({'role': 'user', 'content': message})
            if len(self._sessions[session_id]) > self.MAX_SESSION_HISTORY:
                self._sessions[session_id] = self._sessions[session_id][-self.MAX_SESSION_HISTORY:]
        
        # 1. Check greetings
        greeting = self._check_greeting(message)
        if greeting:
            return self._respond(greeting, 'greeting', session_id)
        
        # 2. Check FAQ knowledge base
        faq_result = self._match_faq(message)
        if faq_result:
            return self._respond(faq_result['answer'], faq_result['category'], session_id, 
                                suggestions=faq_result.get('related', []))
        
        # 3. Check order tracking
        if re.search(r'(订单|查询|跟踪|物流|快递|发货|什么时候|还没收到)', message):
            return self._respond(
                '请提供您的订单号，我可以帮您查询订单状态。您也可以在"我的订单"页面查看最新物流信息。',
                'order', session_id,
                suggestions=['如何查询订单？', '发货时间是多久？', '如何修改收货地址？']
            )
        
        # 4. Check return/refund
        if re.search(r'(退货|退款|退换|不想要|取消)', message):
            return self._respond(
                '我们支持7天无理由退货。如需退换货，请在"我的订单"中选择对应订单申请退换。\n\n注意事项：\n• 香水类产品开封后恕不退换\n• 定制香水不支持无理由退货\n• 退款将在3-7个工作日内退回原支付方式',
                'order', session_id,
                suggestions=['退货流程是什么？', '退款多久到账？', '定制香水能退吗？']
            )
        
        # 5. Check customization questions
        if re.search(r'(定制|调配|香调|配方|前调|中调|后调)', message):
            return self._respond(
                '我们提供专业的香水定制服务！您可以：\n\n1. 前往"香水定制"页面选择香调组合\n2. 参加"香氛测试"获取AI推荐配方\n3. 联系我们的调香师获取专业建议\n\n定制周期通常为3-5个工作日。',
                'customize', session_id,
                suggestions=['定制流程是怎样的？', '如何选择香调？', '定制香水多少钱？']
            )
        
        # 6. Check membership/points
        if re.search(r'(会员|积分|等级|优惠|折扣)', message):
            return self._respond(
                '我们的会员体系包含：\n\n🪙 银卡会员：消费满3000元解锁，享95折\n🥇 金卡会员：消费满10000元解锁，享9折\n💎 钻石会员：消费满30000元解锁，享85折\n\n积分规则：每消费1元=1积分，积分可兑换优惠券和小样。',
                'account', session_id,
                suggestions=['如何获取积分？', '会员有什么权益？', '积分会过期吗？']
            )
        
        # 7. Fallback - needs human
        return self._respond(
            '感谢您的咨询！这个问题比较复杂，建议您联系在线客服获取更详细的帮助。\n\n📞 客服电话：400-888-8888\n📧 邮箱：contact@perfumeshop.com\n🕐 工作时间：周一至周日 9:00-21:00',
            'fallback', session_id,
            needs_human=True,
            suggestions=['联系客服', '发送邮件', '查看帮助中心']
        )
    
    def _check_greeting(self, message: str) -> Optional[str]:
        """Check if message is a greeting."""
        greetings = [
            (r'^(你好|您好|hi|hello|hey|嗨|哈喽)', 
             '您好！👋 我是香氛定制的智能客服小香。\n\n我可以帮您解答：\n• 订单查询与物流跟踪\n• 退换货政策\n• 香水定制流程\n• 会员与积分\n• 产品推荐\n\n请问有什么可以帮您的？'),
            (r'(谢谢|感谢|thanks|thank)', 
             '不客气！😊 如果还有其他问题，随时找我哦~'),
            (r'(再见|拜拜|bye|晚安)', 
             '再见！祝您有美好的一天 🌸\n如有需要随时回来找我~'),
        ]
        for pattern, reply in greetings:
            if re.search(pattern, message, re.IGNORECASE):
                return reply
        return None
    
    def _match_faq(self, message: str) -> Optional[Dict]:
        """Match message against FAQ knowledge base."""
        best_match = None
        best_score = 0
        
        for category, entries in self._faq.items():
            for entry in entries:
                score = self._keyword_match_score(message, entry.get('keywords', []))
                if score > best_score:
                    best_score = score
                    best_match = {
                        'category': category,
                        'answer': entry['answer'],
                        'related': entry.get('related', [])
                    }
        
        # Threshold for matching
        if best_score >= 0.3:
            return best_match
        return None
    
    def _keyword_match_score(self, message: str, keywords: List[str]) -> float:
        """Calculate keyword match score."""
        if not keywords:
            return 0.0
        matches = sum(1 for kw in keywords if kw in message)
        return matches / len(keywords)
    
    def _respond(self, reply: str, category: str, session_id: str = '', 
                 suggestions: List[str] = None, needs_human: bool = False) -> Dict:
        """Build response object."""
        if session_id:
            self._sessions[session_id].append({'role': 'bot', 'content': reply})
        
        return {
            'reply': reply,
            'category': category,
            'needs_human': needs_human,
            'suggestions': suggestions or [],
        }
    
    def _build_faq(self) -> Dict[str, List[Dict]]:
        """Build FAQ knowledge base."""
        return {
            'order': [
                {
                    'question': '如何查询订单状态？',
                    'keywords': ['订单', '查询', '状态', '跟踪'],
                    'answer': '您可以在"我的订单"页面查看所有订单状态。订单状态包括：待付款 → 已支付 → 生产中 → 已发货 → 已完成。点击订单号可查看详细信息。',
                    'related': ['发货时间是多久？', '如何取消订单？', '修改收货地址']
                },
                {
                    'question': '发货时间是多久？',
                    'keywords': ['发货', '多久', '时间', '什么时候'],
                    'answer': '常规商品付款后1-2个工作日发货，定制香水需要3-5个工作日调配制香后发货。全国包邮，顺丰速运配送。',
                    'related': ['如何查询物流？', '可以加急吗？', '配送范围']
                },
                {
                    'question': '如何取消订单？',
                    'keywords': ['取消', '订单', '不想要'],
                    'answer': '在订单状态为"待付款"或"已支付"时，您可以联系客服取消订单。定制香水进入生产环节后无法取消，请谅解。',
                    'related': ['退款多久到账？', '退货流程']
                },
            ],
            'shipping': [
                {
                    'question': '配送范围和费用？',
                    'keywords': ['配送', '包邮', '运费', '范围'],
                    'answer': '全国包邮（港澳台及偏远地区除外），使用顺丰速运配送。满299元免运费，未满299元收取15元运费。',
                },
                {
                    'question': '如何修改收货地址？',
                    'keywords': ['修改', '地址', '收货', '改'],
                    'answer': '在订单未发货前，您可以在"我的订单"中修改收货地址。如已发货，请联系客服协助处理。',
                },
            ],
            'product': [
                {
                    'question': '香水能保存多久？',
                    'keywords': ['保存', '保质期', '过期', '多久'],
                    'answer': '我们的香水准确保质期为3年（未开封）。开封后建议12-24个月内使用完毕。请存放于阴凉干燥处，避免阳光直射。',
                },
                {
                    'question': '如何选择适合自己的香水？',
                    'keywords': ['选择', '适合', '推荐', '怎么选'],
                    'answer': '建议您参加我们的"香氛测试"，通过6道简单的问题即可获得AI个性化香调推荐。也可以浏览产品页面按香调筛选。',
                    'related': ['香氛测试', '热门推荐', '新品上市']
                },
            ],
            'customize': [
                {
                    'question': '定制流程是怎样的？',
                    'keywords': ['定制', '流程', '步骤', '怎么定制'],
                    'answer': '定制流程：1) 选择香调组合（前调/中调/后调）→ 2) 选择容量和瓶型 → 3) 提交定制需求 → 4) 调香师配制 → 5) 发货配送。全程约3-5个工作日。',
                    'related': ['定制香水多少钱？', '可以自己调配吗？', '不满意可以重调吗？']
                },
                {
                    'question': '定制香水多少钱？',
                    'keywords': ['定制', '价格', '多少钱', '费用'],
                    'answer': '定制香水价格取决于容量和瓶型：30ml ¥299起、50ml ¥499起、100ml ¥899起。特殊瓶型和原料可能会有额外费用。',
                },
            ],
            'account': [
                {
                    'question': '如何获取积分？',
                    'keywords': ['积分', '获取', '怎么得', '赚'],
                    'answer': '积分获取方式：消费1元=1积分、每日签到+5积分、发表评价+20积分、分享产品+10积分。积分有效期12个月。',
                    'related': ['积分能做什么？', '积分会过期吗？']
                },
                {
                    'question': '会员等级和权益？',
                    'keywords': ['会员', '等级', '权益', '银卡', '金卡', '钻石'],
                    'answer': '银卡(消费3000+)享95折、金卡(10000+)享9折+生日礼、钻石(30000+)享85折+专属客服+优先发货。',
                    'related': ['如何升级会员？', '会员折扣能叠加吗？']
                },
            ],
        }
