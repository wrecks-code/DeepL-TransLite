from deepl_pptx_translator import config_handler


def add_plus(text) -> str:
    return config_handler.marker_char + text + config_handler.marker_char


def split_text_with_marker(text) -> list:
    # print("splitting " + text)
    segments = [
        segment for segment in text.split(config_handler.marker_char) if segment
    ]
    # print("splitted version: " + str(segments))
    return segments


def assign_segments_to_runs(paragraph, segments):
    # print("size of paragraph.runs: " + str(len(paragraph.runs)))
    # print("size of segments: " + str(len(segments)))

    # Extend the segments list with empty strings if it's shorter than paragraph.runs
    if len(paragraph.runs) > len(segments):
        segments.extend([""] * (len(paragraph.runs) - len(segments)))

    # Iterate through the runs in the paragraph and assign segments to each run
    for i, run in enumerate(paragraph.runs):
        run.text = segments[i]
