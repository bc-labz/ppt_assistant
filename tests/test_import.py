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


if __name__ == "__main__":
    test_import()
    test_class_attributes()
    print("All tests passed!")