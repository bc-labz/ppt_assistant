"""Test basic import functionality."""

def test_import():
    """Test that the package can be imported successfully."""
    try:
        from ppt_assistant import PPTAssistant
        assert PPTAssistant is not None
        print("✓ Package import successful")
    except ImportError as e:
        raise AssertionError(f"Failed to import package: {e}")


def test_class_attributes():
    """Test that the PPTAssistant class and config work correctly."""
    from ppt_assistant import PPTAssistant, DEFAULT_CONFIG

    # Check that config has expected attributes
    assert hasattr(DEFAULT_CONFIG, 'SLIDE_WIDTH')
    assert hasattr(DEFAULT_CONFIG, 'SLIDE_HEIGHT')
    assert DEFAULT_CONFIG.SLIDE_WIDTH == 13.333
    assert DEFAULT_CONFIG.SLIDE_HEIGHT == 7.5

    # Test that PPTAssistant accepts config parameter
    assert hasattr(PPTAssistant.__init__, '__code__')
    init_params = PPTAssistant.__init__.__code__.co_varnames
    assert 'ppt_config' in init_params

    print("✓ Class and config attributes verified")


def test_config_validation():
    """Test that config validation works correctly."""
    from ppt_assistant import PPTConfig

    # Valid config should work
    config = PPTConfig()
    assert config.content_width > 0
    assert config.content_height > 0

    # Test invalid slide width
    try:
        config.SLIDE_WIDTH = 0.5  # Too small
        config._validate_config()
        assert False, "Should have raised ValueError"
    except ValueError:
        pass  # Expected

    # Test invalid ratios
    try:
        config = PPTConfig()
        config.IMAGE_WIDTH_RATIO = 1.5  # > 1
        config._validate_config()
        assert False, "Should have raised ValueError"
    except ValueError:
        pass  # Expected

    # Test invalid color values
    try:
        config = PPTConfig()
        config.TABLE_HEADER_BG_COLOR = (300, 100, 100)  # Value > 255
        config._validate_config()
        assert False, "Should have raised ValueError"
    except ValueError:
        pass  # Expected

    print("✓ Config validation verified")


def test_config_calculated_properties():
    """Test that calculated properties work correctly."""
    from ppt_assistant import PPTConfig

    config = PPTConfig()

    # Test calculated dimensions
    expected_content_width = config.SLIDE_WIDTH - (2 * config.MARGIN)
    assert config.content_width == expected_content_width

    expected_content_top = config.TITLE_HEIGHT + config.CONTENT_TOP_OFFSET
    assert config.content_top == expected_content_top

    expected_content_height = config.SLIDE_HEIGHT - config.content_top - config.MARGIN
    assert config.content_height == expected_content_height

    print("✓ Config calculated properties verified")


def test_custom_config_application():
    """Test that custom configurations are properly applied."""
    from ppt_assistant import PPTAssistant, PPTConfig
    import tempfile
    import os

    # Create a custom config
    custom_config = PPTConfig()
    custom_config.TITLE_FONT_SIZE = 48
    custom_config.SLIDE_WIDTH = 10.0  # Smaller slide

    # Create a temporary JSON file
    test_json = {
        "template": "examples/template.pptx",
        "output": "test_output.pptx",
        "slides": [
            {
                "layout": "title",
                "content": {
                    "title": "Test Title",
                    "subtitle": "Test Subtitle"
                }
            }
        ]
    }

    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        import json
        json.dump(test_json, f)
        temp_json_path = f.name

    try:
        # Test with custom config
        assistant = PPTAssistant(temp_json_path, custom_config)

        # Verify config is used (check that custom slide width is applied)
        # We can't easily test font sizes without generating the presentation,
        # but we can verify the config is stored
        assert assistant.ppt_config.TITLE_FONT_SIZE == 48
        assert assistant.ppt_config.SLIDE_WIDTH == 10.0

        print("✓ Custom config application verified")

    finally:
        # Clean up
        if os.path.exists(temp_json_path):
            os.unlink(temp_json_path)
        if os.path.exists("test_output.pptx"):
            os.unlink("test_output.pptx")


if __name__ == "__main__":
    test_import()
    test_class_attributes()
    print("All tests passed!")