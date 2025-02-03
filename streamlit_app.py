
import os
from pptx import Presentation
from openpyxl import Workbook

def analyze_slide(slide, criteria):
    """Analyze the slide based on given criteria."""
    feedback = {}
    
    # 1. Clarity of Content (Bullet Points and Text)
    if "bullet_points" in criteria:
        bullet_points = [shape.text for shape in slide.shapes if shape.shape_type == 3]
        feedback["clarity"] = f"{len(bullet_points)} bullet points found."
        
    if "text_length" in criteria:
        text_shapes = []
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                text_shapes.append(len(shape.text.split()))
        total_words = sum(text_shapes)
        feedback["clarity"] += f" Total words: {total_words}."
    
    # 2. Strategy (Slide Position and Messaging)
    if "slide_position" in criteria:
        slide_number = slide.slide_number
        feedback["strategy"] = f"Slide {slide_number}: Ensure messaging aligns with overall strategy."
        
    # 3. Outcome Focus
    if "outcome_focus" in criteria:
        has_outcome = any("Outcome:" in shape.text for shape in slide.shapes)
        feedback["outcome"] = "Focus on outcomes." + (" Yes, outcome identified." if has_outcome else ". No outcome identified.")
    
    # 4. Descriptors and Data
    if "descriptors" in criteria:
        charts_present = sum(1 for shape in slide.shapes if hasattr(shape.chart, 'chart_type'))
        feedback["descriptors"] = f"{charts_present} charts found. Ensure data is clearly described."
        
    # 5. Analyses and Insights
    if "analyses" in criteria:
        analysis_shapes = [shape.text for shape in slide.shapes if any(word in shape.text.lower() for word in ["analyze", "insight"])]
        feedback["analysis"] = f"{len(analysis_shapes)} insights identified."
    
    # 6. Recommendations and Next Steps
    if "recommendations" in criteria:
        rec_shapes = [shape.text for shape in slide.shapes if any(word in shape.text.lower() for word in ["recommend", "next step"])]
        feedback["recommendation"] = f"{len(rec_shapes)} recommendations found."
    
    return feedback

def generate_feedback_report(ppt_path):
    """Generate a feedback report based on the PowerPoint analysis."""
    prs = Presentation(ppt_path)
    wb = Workbook()
    ws = wb.active
    ws.title = "Slide Feedback"
    
    # Create headers in Excel sheet
    headers = ["Slide Number", 
               "Clarity of Content",
               "Strategy Alignment",
               "Outcome Focus",
               "Data Descriptors",
               "Analyses and Insights",
               "Recommendations"]
    ws.append(headers)
    
    for slide_num, slide in enumerate(prs.slides):
        feedback = analyze_slide(slide, criteria=["bullet_points", 
                                                  "text_length", 
                                                  "slide_position", 
                                                  "outcome_focus", 
                                                  "descriptors", 
                                                  "analyses",
                                                  "recommendations"])
        
        row_data = [
            f"Slide {slide_num + 1}",
            feedback.get("clarity", ""),
            feedback.get("strategy", ""),
            feedback.get("outcome", ""),
            feedback.get("descriptors", ""),
            feedback.get("analysis", ""),
            feedback.get("recommendation", "")
        ]
        ws.append(row_data)
    
    # Save the Excel file
    wb.save("feedback.xlsx")
    print(f"Feedback saved to 'feedback.xlsx'.")

# Example usage:
if __name__ == "__main__":
    ppt_path = input("Enter the path of your PowerPoint file: ")
    if os.path.exists(ppt_path):
        generate_feedback_report(ppt_path)
        print("Analysis complete. Check feedback.xlsx for results.")
    else:
        print(f"File not found at {ppt_path}.")


