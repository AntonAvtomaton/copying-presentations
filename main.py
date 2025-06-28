import find_dirs
import func_to_copy

for pres in find_dirs.Presentations:
    func_to_copy.copy_slide_to_new_presentation(
        source_file_dir=pres,
        output_file=find_dirs.New_pres,
    )

