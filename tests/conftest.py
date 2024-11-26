def pytest_configure(config):
    config.addinivalue_line("markers", "description: תיאור לבדיקות")

import pytest

@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    outcome = yield
    report = outcome.get_result()
    if report.when == 'call':
        feature_request = item.funcargs['request']
        if hasattr(feature_request.node, "extra"):
            report.extras = feature_request.node.extra


terms = ["ביטוח","ברוש"]
