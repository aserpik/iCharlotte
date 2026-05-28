from icharlotte_core.ui.wizard.registry import TaskSpec


def test_taskspec_default_output_page_cls_is_OutputPage():
    spec = TaskSpec(task_id="x", title="X", description="x",
                    icon_glyph="X", script_name="x.py")
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage
    assert spec.output_page_cls is OutputPage


def test_taskspec_custom_output_factory_is_used():
    class MyOutputPage:
        pass
    spec = TaskSpec(
        task_id="x", title="X", description="x", icon_glyph="X", script_name="x.py",
        _output_page_cls_factory=lambda: MyOutputPage,
    )
    assert spec.output_page_cls is MyOutputPage
