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
    """Test that the PPTAssistant class has expected attributes."""
    from ppt_assistant import PPTAssistant

    # Check class constants
    assert hasattr(PPTAssistant, 'SLIDE_WIDTH')
    assert hasattr(PPTAssistant, 'SLIDE_HEIGHT')
    assert PPTAssistant.SLIDE_WIDTH == 13.333
    assert PPTAssistant.SLIDE_HEIGHT == 7.5

    print("✓ Class attributes verified")


if __name__ == "__main__":
    test_import()
    test_class_attributes()
    print("All tests passed!")