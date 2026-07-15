import threading
import unittest
from unittest.mock import patch

import web_query


class FakeResponse:
    def __init__(self, status=200):
        self.status = status
        self.ok = 200 <= status < 300


class FakePage:
    def __init__(
        self,
        name,
        events=None,
        goto_error=None,
        redirected_url=None,
        response_status=200,
    ):
        self.name = name
        self.events = events if events is not None else []
        self.goto_error = goto_error
        self.redirected_url = redirected_url
        self.response_status = response_status
        self.closed = False
        self.close_options = []
        self.brought_to_front = False
        self.navigations = []
        self.url = f"https://example.test/{name}"
        self._handlers = {"close": []}
        self.after_close = None

    def is_closed(self):
        return self.closed

    def close(self, **kwargs):
        self.close_options.append(kwargs)
        self.closed = True
        self.events.append(f"close:{self.name}")
        for handler in self._handlers["close"]:
            handler(self)
        if self.after_close:
            self.after_close()

    def bring_to_front(self):
        self.brought_to_front = True

    def goto(self, url, **kwargs):
        self.navigations.append((url, kwargs))
        self.events.append(f"goto:{self.name}")
        if self.goto_error:
            raise self.goto_error
        self.url = self.redirected_url or url
        return FakeResponse(self.response_status)

    def on(self, event, handler):
        self._handlers.setdefault(event, []).append(handler)


class FakeDeferredClosePage(FakePage):
    """模拟已收到关闭请求、但 Edge 尚未完成关闭的后台页。"""

    def close(self, **kwargs):
        self.close_options.append(kwargs)
        self.events.append(f"close-request:{self.name}")


class FakeContext:
    def __init__(self, *pages, fallback_page=None):
        self.pages = list(pages)
        self.new_page_calls = 0
        self.cdp_sessions = []
        self.fallback_page = fallback_page

    def new_page(self):
        self.new_page_calls += 1
        page = self.fallback_page or FakePage("fallback")
        self.pages.append(page)
        return page

    def new_cdp_session(self, page):
        session = FakeCDPSession(page)
        self.cdp_sessions.append(session)
        return session


class FakePageExpectation:
    def __init__(self, context):
        self.context = context
        self.value = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if exc_type is not None:
            return False
        if len(self.context.pages) <= 1:
            raise RuntimeError("no popup opened")
        self.value = self.context.pages[-1]
        return False


class FakePopupContext(FakeContext):
    def expect_page(self, timeout):
        self.expect_page_timeout = timeout
        return FakePageExpectation(self)


class FakeRequestElement:
    def __init__(self, context, detail_page):
        self.context = context
        self.detail_page = detail_page
        self.actions = []

    def is_visible(self, timeout):
        return True

    def inner_text(self):
        return "LM8000001"

    def click(self):
        self.actions.append("click")
        self.context.pages.append(self.detail_page)

    def evaluate(self, expression):
        self.actions.append(("evaluate", expression))
        raise RuntimeError("JS click must not be used")


class FakeDelayedRequestElement(FakeRequestElement):
    """请求已发出，但详情页会在延迟等待阶段才出现在上下文中。"""

    def click(self):
        self.actions.append("click")


class FakeSearchPageForRequestClick(FakePage):
    def __init__(self, request_element):
        super().__init__("search")
        self.url = web_query.SEARCH_URL
        self.request_element = request_element

    def locator(self, selector):
        return self

    @property
    def first(self):
        return self.request_element


class FakeCDPSession:
    def __init__(self, page):
        self.page = page
        self.commands = []
        self.detached = False

    def send(self, method, params=None):
        self.commands.append((method, params))

    def detach(self):
        self.detached = True


class SearchPageRotationTests(unittest.TestCase):
    def setUp(self):
        web_query._PAGE_ACTIVITY_SESSIONS.clear()

    def test_edge_launch_configuration_prevents_background_page_throttling(self):
        self.assertTrue(
            {
                "--disable-background-timer-throttling",
                "--disable-renderer-backgrounding",
                "--disable-backgrounding-occluded-windows",
            }.issubset(web_query.EDGE_BACKGROUND_PROTECTION_ARGS)
        )

    def test_query_can_stop_before_opening_the_browser(self):
        stop_event = threading.Event()
        stop_event.set()

        success, message = web_query.query_and_download_attachments(
            stop_requested=stop_event,
        )

        self.assertTrue(success)
        self.assertIn("已按用户请求停止", message)

    def test_restores_page_focus_and_active_lifecycle_before_navigation(self):
        detail_page = FakePage("active-detail")
        context = FakeContext(detail_page)
        messages = []

        web_query._restore_page_activity(context, detail_page, messages.append)

        self.assertTrue(detail_page.brought_to_front)
        self.assertEqual(len(context.cdp_sessions), 1)
        self.assertEqual(
            context.cdp_sessions[0].commands,
            [
                ("Emulation.setFocusEmulationEnabled", {"enabled": True}),
                ("Page.setWebLifecycleState", {"state": "active"}),
            ],
        )
        # 焦点模拟是会话级状态；会话需要保持到页面关闭，不能立刻 detach。
        self.assertFalse(context.cdp_sessions[0].detached)
        self.assertTrue(any("页面活动状态已恢复" in message for message in messages))

    def test_releases_activity_session_when_the_page_is_closed(self):
        detail_page = FakePage("active-detail")
        context = FakeContext(detail_page)

        web_query._restore_page_activity(context, detail_page, lambda _message: None)
        web_query._close_other_pages(context, FakePage("keep"), lambda _message: None)

        self.assertTrue(detail_page.closed)
        self.assertTrue(context.cdp_sessions[0].detached)

    def test_releases_activity_session_when_page_closes_outside_cleanup(self):
        detail_page = FakePage("active-detail")
        context = FakeContext(detail_page)

        web_query._restore_page_activity(context, detail_page, lambda _message: None)
        detail_page.close()

        self.assertNotIn(detail_page, web_query._PAGE_ACTIVITY_SESSIONS)

    def test_reuses_active_detail_page_for_the_next_search(self):
        old_search_page = FakePage("old-search")
        detail_page = FakePage("active-detail")
        context = FakeContext(old_search_page, detail_page)
        messages = []

        next_search_page = web_query._prepare_next_search_page(
            context, old_search_page, detail_page, messages.append
        )

        self.assertIs(next_search_page, detail_page)
        self.assertTrue(old_search_page.closed)
        self.assertFalse(detail_page.closed)
        self.assertTrue(detail_page.brought_to_front)
        self.assertEqual(
            detail_page.navigations,
            [(web_query.SEARCH_URL, {"wait_until": "domcontentloaded", "timeout": 20000})],
        )
        self.assertEqual(context.new_page_calls, 0)

    def test_navigates_the_active_detail_page_before_closing_the_old_search_page(self):
        events = []
        old_search_page = FakePage("old-search", events)
        detail_page = FakePage("active-detail", events)
        context = FakeContext(old_search_page, detail_page)

        web_query._prepare_next_search_page(
            context, old_search_page, detail_page, lambda _message: None
        )

        self.assertLess(
            events.index("goto:active-detail"),
            events.index("close:old-search"),
        )

    def test_returns_same_page_detail_to_search_and_cleans_orphans(self):
        search_and_detail_page = FakePage("same-page")
        orphan_detail_page = FakePage("orphan-detail")
        context = FakeContext(search_and_detail_page, orphan_detail_page)

        next_search_page = web_query._prepare_next_search_page(
            context,
            search_and_detail_page,
            search_and_detail_page,
            lambda _message: None,
        )

        self.assertIs(next_search_page, search_and_detail_page)
        self.assertEqual(search_and_detail_page.url, web_query.SEARCH_URL)
        self.assertTrue(orphan_detail_page.closed)

    def test_keeps_old_search_page_when_fallback_page_cannot_load(self):
        old_search_page = FakePage("old-search")
        detail_page = FakePage("active-detail", goto_error=RuntimeError("detail goto failed"))
        fallback_page = FakePage("fallback", goto_error=RuntimeError("fallback goto failed"))
        context = FakeContext(old_search_page, detail_page, fallback_page=fallback_page)

        next_search_page = web_query._prepare_next_search_page(
            context, old_search_page, detail_page, lambda _message: None
        )

        self.assertIs(next_search_page, old_search_page)
        self.assertFalse(old_search_page.closed)
        self.assertTrue(fallback_page.closed)

    def test_keeps_old_search_page_when_fallback_redirects_elsewhere(self):
        old_search_page = FakePage("old-search")
        detail_page = FakePage("active-detail", goto_error=RuntimeError("detail goto failed"))
        fallback_page = FakePage(
            "fallback", redirected_url="https://view.tkelevator.com.cn/login.php"
        )
        context = FakeContext(old_search_page, detail_page, fallback_page=fallback_page)

        next_search_page = web_query._prepare_next_search_page(
            context, old_search_page, detail_page, lambda _message: None
        )

        self.assertIs(next_search_page, old_search_page)
        self.assertFalse(old_search_page.closed)
        self.assertTrue(fallback_page.closed)

    def test_keeps_old_search_page_when_fallback_adds_an_error_query(self):
        old_search_page = FakePage("old-search")
        detail_page = FakePage("active-detail", goto_error=RuntimeError("detail goto failed"))
        fallback_page = FakePage(
            "fallback", redirected_url=f"{web_query.SEARCH_URL}?error=session-expired"
        )
        context = FakeContext(old_search_page, detail_page, fallback_page=fallback_page)

        next_search_page = web_query._prepare_next_search_page(
            context, old_search_page, detail_page, lambda _message: None
        )

        self.assertIs(next_search_page, old_search_page)
        self.assertFalse(old_search_page.closed)
        self.assertTrue(fallback_page.closed)

    def test_keeps_old_search_page_when_fallback_returns_http_error(self):
        old_search_page = FakePage("old-search")
        detail_page = FakePage("active-detail", goto_error=RuntimeError("detail goto failed"))
        fallback_page = FakePage("fallback", response_status=500)
        context = FakeContext(old_search_page, detail_page, fallback_page=fallback_page)

        next_search_page = web_query._prepare_next_search_page(
            context, old_search_page, detail_page, lambda _message: None
        )

        self.assertIs(next_search_page, old_search_page)
        self.assertFalse(old_search_page.closed)
        self.assertTrue(fallback_page.closed)

    def test_uses_an_open_verified_search_page_when_detail_page_is_missing(self):
        stale_search_page = FakePage("stale-search")
        stale_search_page.closed = True
        verified_search_page = FakePage("verified-search")
        verified_search_page.url = web_query.SEARCH_URL
        orphan_detail_page = FakePage("orphan-detail")
        context = FakeContext(stale_search_page, verified_search_page, orphan_detail_page)

        next_search_page = web_query._prepare_next_search_page(
            context, stale_search_page, None, lambda _message: None
        )

        self.assertIs(next_search_page, verified_search_page)
        self.assertTrue(orphan_detail_page.closed)

    def test_closes_untracked_detail_pages_after_rotation(self):
        old_search_page = FakePage("old-search")
        detail_page = FakePage("active-detail")
        orphan_detail_page = FakePage("orphan-detail")
        context = FakeContext(old_search_page, detail_page, orphan_detail_page)

        next_search_page = web_query._prepare_next_search_page(
            context, old_search_page, detail_page, lambda _message: None
        )

        self.assertIs(next_search_page, detail_page)
        self.assertTrue(old_search_page.closed)
        self.assertTrue(orphan_detail_page.closed)
        self.assertFalse(detail_page.closed)

    def test_closes_page_that_opens_while_an_orphan_is_being_closed(self):
        keep_page = FakePage("keep")
        orphan_detail_page = FakePage("orphan-detail")
        context = FakeContext(keep_page, orphan_detail_page)
        late_popup_pages = []

        def open_late_popup():
            late_popup = FakePage("late-popup")
            late_popup_pages.append(late_popup)
            context.pages.append(late_popup)

        orphan_detail_page.after_close = open_late_popup

        all_pages_closed = web_query._close_other_pages(
            context, keep_page, lambda _message: None
        )

        self.assertTrue(all_pages_closed)
        self.assertTrue(late_popup_pages[0].closed)

    def test_cleanup_requests_nonblocking_close_for_stale_pages(self):
        keep_page = FakePage("keep")
        stale_page = FakePage("stale")
        context = FakeContext(keep_page, stale_page)

        web_query._close_other_pages(context, keep_page, lambda _message: None)

        self.assertEqual(stale_page.close_options, [{"run_before_unload": True}])

    def test_cleanup_does_not_repeat_a_pending_close_request(self):
        keep_page = FakePage("keep")
        stale_page = FakeDeferredClosePage("stale")
        context = FakeContext(keep_page, stale_page)
        messages = []

        all_pages_closed = web_query._close_other_pages(
            context, keep_page, messages.append
        )

        self.assertFalse(all_pages_closed)
        self.assertEqual(stale_page.close_options, [{"run_before_unload": True}])
        self.assertTrue(any("不阻塞下一 PO" in message for message in messages))

    def test_clicks_request_id_once_while_waiting_for_its_popup(self):
        detail_page = FakePage("detail")
        context = FakePopupContext()
        request_element = FakeRequestElement(context, detail_page)
        search_page = FakeSearchPageForRequestClick(request_element)
        context.pages = [search_page]

        with patch.object(web_query.time, "sleep"):
            request_id, opened_page = web_query._click_request_id(
                search_page, context, lambda _message: None
            )

        self.assertEqual(request_id, "LM8000001")
        self.assertIs(opened_page, detail_page)
        self.assertEqual(
            request_element.actions,
            ["click"],
        )
        self.assertEqual(context.expect_page_timeout, 12000)

    def test_uses_one_trusted_click_when_js_click_does_not_trigger_navigation(self):
        """少数请求 ID 仅响应浏览器真实点击，JS click 不应成为唯一途径。"""
        class TrustedClickRequestElement(FakeRequestElement):
            def click(self):
                self.actions.append("click")
                self.context.pages.append(self.detail_page)

            def evaluate(self, expression):
                self.actions.append(("evaluate", expression))

        detail_page = FakePage("trusted-click-detail")
        context = FakePopupContext()
        request_element = TrustedClickRequestElement(context, detail_page)
        search_page = FakeSearchPageForRequestClick(request_element)
        context.pages = [search_page]

        with patch.object(web_query.time, "sleep"):
            request_id, opened_page = web_query._click_request_id(
                search_page, context, lambda _message: None
            )

        self.assertEqual(request_id, "LM8000001")
        self.assertIs(opened_page, detail_page)
        self.assertEqual(request_element.actions, ["click"])

    def test_waits_for_a_late_detail_page_without_clicking_the_request_id_again(self):
        detail_page = FakePage("late-detail")
        context = FakePopupContext()
        request_element = FakeDelayedRequestElement(context, detail_page)
        search_page = FakeSearchPageForRequestClick(request_element)
        context.pages = [search_page]
        sleep_calls = []

        def delayed_sleep(_seconds):
            sleep_calls.append(_seconds)
            if len(sleep_calls) == 5:
                context.pages.append(detail_page)

        with patch.object(web_query.time, "sleep", side_effect=delayed_sleep):
            request_id, opened_page = web_query._click_request_id(
                search_page, context, lambda _message: None
            )

        self.assertEqual(request_id, "LM8000001")
        self.assertIs(opened_page, detail_page)
        self.assertEqual(
            request_element.actions,
            ["click"],
        )

    def test_keeps_observing_when_detail_page_opens_after_the_old_deadline(self):
        """站点可能在首次 12 秒等待与原 15 秒补等之后才真正创建详情页。"""
        detail_page = FakePage("very-late-detail")
        context = FakePopupContext()
        request_element = FakeDelayedRequestElement(context, detail_page)
        search_page = FakeSearchPageForRequestClick(request_element)
        context.pages = [search_page]
        sleep_calls = []

        def delayed_sleep(_seconds):
            sleep_calls.append(_seconds)
            # 详情页在原来的 15 秒补等结束后才出现，不能被清理成“多余页面”。
            if len(sleep_calls) == 20:
                context.pages.append(detail_page)

        with patch.object(web_query.time, "sleep", side_effect=delayed_sleep):
            request_id, opened_page = web_query._click_request_id(
                search_page, context, lambda _message: None
            )

        self.assertEqual(request_id, "LM8000001")
        self.assertIs(opened_page, detail_page)
        self.assertEqual(
            request_element.actions,
            ["click"],
        )

    def test_keeps_observing_for_a_detail_page_that_appears_after_two_minutes(self):
        """极慢站点响应不能在 PO 切换前被误判为失败并清理掉详情页。"""
        detail_page = FakePage("two-minute-late-detail")
        context = FakePopupContext()
        request_element = FakeDelayedRequestElement(context, detail_page)
        search_page = FakeSearchPageForRequestClick(request_element)
        context.pages = [search_page]
        sleep_calls = []

        def delayed_sleep(_seconds):
            sleep_calls.append(_seconds)
            # 公司日志证明：点击后约 118 秒，详情页才实际出现在浏览器上下文中。
            if len(sleep_calls) == 118:
                context.pages.append(detail_page)

        with patch.object(web_query.time, "sleep", side_effect=delayed_sleep):
            request_id, opened_page = web_query._click_request_id(
                search_page, context, lambda _message: None
            )

        self.assertEqual(request_id, "LM8000001")
        self.assertIs(opened_page, detail_page)
        self.assertEqual(request_element.actions, ["click"])

    def test_opens_a_controlled_page_when_visible_detail_is_missing_from_context(self):
        """可见详情页未注册到 context.pages 时，仍须按浏览器目标 URL 接管它。"""
        class BrowserTargetSession:
            def __init__(self, target_url):
                self.target_url = target_url
                self.commands = []
                self.detached = False

            def send(self, method, params=None):
                self.commands.append((method, params))
                if method == "Target.getTargets":
                    return {
                        "targetInfos": [
                            {
                                "targetId": "untracked-detail",
                                "type": "page",
                                "url": self.target_url,
                            }
                        ]
                    }
                return {}

            def detach(self):
                self.detached = True

        class BrowserWithTargets:
            def __init__(self, target_url):
                self.session = BrowserTargetSession(target_url)

            def new_browser_cdp_session(self):
                return self.session

        target_url = (
            "https://view.tkelevator.com.cn/vivid/niops/purchasing/materials/"
            "requests/approval?projectId=1&branchId=2&requestId=8000001"
        )
        controlled_detail = FakePage("controlled-detail")
        context = FakePopupContext(fallback_page=controlled_detail)
        context.browser = BrowserWithTargets(target_url)
        request_element = FakeDelayedRequestElement(context, FakePage("ignored"))
        search_page = FakeSearchPageForRequestClick(request_element)
        context.pages = [search_page]

        with patch.object(web_query.time, "sleep"):
            request_id, opened_page = web_query._click_request_id(
                search_page, context, lambda _message: None
            )

        self.assertEqual(request_id, "LM8000001")
        self.assertIs(opened_page, controlled_detail)
        self.assertEqual(controlled_detail.navigations[0][0], target_url)
        self.assertEqual(request_element.actions, ["click"])
        self.assertIn(
            ("Target.closeTarget", {"targetId": "untracked-detail"}),
            context.browser.session.commands,
        )
        self.assertTrue(context.browser.session.detached)

    def test_recognizes_current_request_when_site_reuses_an_existing_detail_tab(self):
        """详情可能复用旧标签，不能只依赖“点击后新建页面”的事件。"""
        class NoPopupExpectation:
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc_value, traceback):
                if exc_type is not None:
                    return False
                raise RuntimeError("no newly created popup")

        class ExistingTabContext(FakeContext):
            def expect_page(self, timeout):
                self.expect_page_timeout = timeout
                return NoPopupExpectation()

        class ExistingTabRequestElement(FakeDelayedRequestElement):
            def click(self):
                self.actions.append("click")
                self.detail_page.url = (
                    "https://view.tkelevator.com.cn/vivid/niops/purchasing/"
                    "materials/requests/approval?projectId=1&requestId=8000001"
                )

        detail_page = FakePage("reused-detail")
        detail_page.url = "about:blank"
        context = ExistingTabContext()
        request_element = ExistingTabRequestElement(context, detail_page)
        search_page = FakeSearchPageForRequestClick(request_element)
        context.pages = [search_page, detail_page]

        with patch.object(web_query.time, "sleep"):
            request_id, opened_page = web_query._click_request_id(
                search_page, context, lambda _message: None
            )

        self.assertEqual(request_id, "LM8000001")
        self.assertIs(opened_page, detail_page)
        self.assertEqual(
            request_element.actions,
            ["click"],
        )

    def test_does_not_report_a_po_as_completed_when_its_detail_page_never_opens(self):
        class FakeKeyboard:
            def press(self, _key):
                pass

            def type(self, _text, **_kwargs):
                pass

        class FakePOInput:
            def click(self, **_kwargs):
                pass

        class FakeSearchPage:
            def __init__(self):
                self.keyboard = FakeKeyboard()

        search_page = FakeSearchPage()
        next_search_page = object()
        logs = []

        with patch.object(web_query, "_restore_page_activity"), patch.object(
            web_query, "_locate_po_input", return_value=FakePOInput()
        ), patch.object(web_query, "_locate_search_button", return_value=None), patch.object(
            web_query, "_wait_for_search_results", return_value=True
        ), patch.object(
            web_query, "_click_request_id", return_value=("LM8000001", None)
        ), patch.object(
            web_query, "_prepare_next_search_page", return_value=next_search_page
        ):
            result, returned_search_page = web_query._process_one_po(
                "4000000001", object(), search_page, logs.append
            )

        self.assertIsNone(result)
        self.assertIs(returned_search_page, next_search_page)
        self.assertTrue(any("未能进入详情页" in message for message in logs))


if __name__ == "__main__":
    unittest.main()
