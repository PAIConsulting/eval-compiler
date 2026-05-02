"""
generate_fake_data.py

Generates synthetic evaluation form responses for testing the compiler tool.

Project:    Evaluation Compiler (InfoShare evaluation reporting tool)
Owner:      PAI Consulting
Author:     Amanda Rummel (with code generation assistance from Claude / Anthropic AI)
Version:    0.1.0 (demonstration)
Created:    2026
License:    Proprietary - see LICENSE file

================================================================================
AUDIT-RELEVANT NOTES
================================================================================

This script:
  - Runs entirely on the local machine
  - Reads NO input files
  - Writes ONE output file: fake_responses.csv in the working directory
  - Does NOT make any network calls
  - Does NOT use any AI, machine-learning, or LLM libraries
  - Does NOT execute external commands or shell processes
  - Uses Python's standard library `random` module (with fixed seed 42 for
    reproducibility) to generate synthetic data
  - All generated data is fictional; no resemblance to real evaluation
    responses is intended

Dependencies:
  - Python standard library only (csv, random)

================================================================================
USAGE
================================================================================

    python generate_fake_data.py

Produces fake_responses.csv with synthetic evaluation responses.
"""

import csv
import random

# ==================== FICTIONAL FORM CONFIGURATION ====================
# All values below are made-up examples for demonstration only.

DOMAINS = [
    "Operations - Track A",
    "Operations - Track B",
    "Maintenance",
]

Q2_TOPICS = [
    "Topic Alpha - Equipment Reliability",
    "Topic Beta - Crew Resource Management",
    "Topic Gamma - Communication Protocols",
    "Topic Delta - Training Standards",
    "Topic Epsilon - Documentation",
    "Topic Zeta - Risk Assessment",
    "Topic Eta - Emergency Procedures",
    "Topic Theta - Maintenance Scheduling",
    "Topic Iota - Personnel Wellness",
    "Topic Kappa - Technology Adoption",
]

SHARED_SUBJECTS = [
    "Subject 1 - Personnel",
    "Subject 2 - Equipment",
    "Subject 3 - Procedures",
    "Subject 4 - Training",
    "Subject 5 - Communication",
    "Subject 6 - External Factors",
    "Other",
]

TOPICS_BY_SUBJECT = {
    "Subject 1 - Personnel": [
        "Fatigue management",
        "Workload distribution",
        "Staffing levels",
        "Experience levels",
        "Other",
    ],
    "Subject 2 - Equipment": [
        "Aging hardware",
        "Software updates",
        "Sensor reliability",
        "Spare parts availability",
        "Other",
    ],
    "Subject 3 - Procedures": [
        "Standard operating procedures",
        "Checklists",
        "Documentation gaps",
        "Process compliance",
        "Other",
    ],
    "Subject 4 - Training": [
        "Initial training quality",
        "Recurrent training frequency",
        "Simulator availability",
        "Cross-training",
        "Other",
    ],
    "Subject 5 - Communication": [
        "Inter-team communication",
        "Reporting channels",
        "Information dissemination",
        "Language barriers",
        "Other",
    ],
    "Subject 6 - External Factors": [
        "Weather impacts",
        "Regulatory changes",
        "Supply chain issues",
        "Public perception",
        "Other",
    ],
    "Other": [
        "Other",
    ],
}

# Sample fake comment phrases - these are pure lorem-ipsum-style fillers,
# not realistic content. The compiler treats them as opaque text to demonstrate
# keyword matching.
FAKE_COMMENT_PHRASES = [
    "We have noticed increased issues related to fatigue among night shift staff.",
    "Equipment failures have caused minor delays but no safety incidents recently.",
    "Training programs need to be updated to reflect new procedures.",
    "Communication between teams could be significantly improved.",
    "Documentation is sometimes unclear and leads to inconsistent practices.",
    "Staffing shortages are creating additional workload pressure.",
    "New software rollout went smoothly but training was minimal.",
    "Weather-related disruptions have been higher than usual this quarter.",
    "We need better tools for tracking maintenance schedules.",
    "Cross-team meetings have been very productive when they occur.",
    "Some checklists are outdated and reference old equipment.",
    "Onboarding for new staff takes longer than expected.",
    "Reporting channels are unclear for non-routine observations.",
    "Spare parts inventory is sometimes insufficient for routine maintenance.",
    "Workload during peak periods is unsustainable without changes.",
    "",  # Empty - simulating respondents who skipped the comment field
    "",
    "",
]

FAKE_OPEN_RESPONSES = {
    "Q6": [
        "Add more interactive sessions to break up the day.",
        "More time for networking would be valuable.",
        "Consider hybrid attendance options for those who can't travel.",
        "The presentation slides could be made available in advance.",
        "Smaller breakout groups would encourage more participation.",
        "Focus more on practical applications and less on theory.",
        "",
        "",
    ],
    "Q7": [
        "Continued staffing pressures and turnover.",
        "Adoption of new technology without adequate training.",
        "Aging workforce and knowledge transfer challenges.",
        "Increasing regulatory complexity.",
        "Supply chain disruptions affecting maintenance.",
        "Changes in passenger volume and patterns.",
        "",
        "",
    ],
    "Q8": [
        "Great event overall, looking forward to next year.",
        "The venue could be improved.",
        "More diverse perspectives in the panels would help.",
        "Logistics were smoothly handled.",
        "Some sessions ran long; tighter scheduling would help.",
        "The catering was excellent.",
        "",
        "",
    ],
}


def make_response(respondent_id):
    """Create one fake response row."""
    domain = random.choice(DOMAINS)
    
    # Q2 - pick 1 to 4 topics
    num_q2 = random.randint(1, 4)
    q2 = random.sample(Q2_TOPICS, num_q2)
    
    # Q3, Q4, Q5 - cascading subject + topic + comment
    def cascading():
        subject = random.choice(SHARED_SUBJECTS)
        topic = random.choice(TOPICS_BY_SUBJECT[subject])
        comment = random.choice(FAKE_COMMENT_PHRASES)
        # If "Other" was chosen, simulate a write-in
        subject_other = ""
        topic_other = ""
        if subject == "Other":
            subject_other = "Custom write-in subject example"
        if topic == "Other":
            topic_other = random.choice([
                "Pilot fatigue concern",
                "Software glitch in console",
                "Training simulator availability",
                "Communication delay",
                "",
            ])
        return subject, subject_other, topic, topic_other, comment
    
    q3_subj, q3_subj_other, q3_topic, q3_topic_other, q3_comment = cascading()
    q4_subj, q4_subj_other, q4_topic, q4_topic_other, q4_comment = cascading()
    q5_subj, q5_subj_other, q5_topic, q5_topic_other, q5_comment = cascading()
    
    # Q6, Q7, Q8 - free text
    q6 = random.choice(FAKE_OPEN_RESPONSES["Q6"])
    q7 = random.choice(FAKE_OPEN_RESPONSES["Q7"])
    q8 = random.choice(FAKE_OPEN_RESPONSES["Q8"])
    
    # Q9 - first time attendee
    q9 = random.choice(["Yes", "No", "No", "No"])  # Most are returning attendees
    
    return {
        "Respondent_ID": respondent_id,
        "Q1_Domain": domain,
        "Q2_Topics": "; ".join(q2),  # multi-select stored semicolon-separated
        "Q3_Subject": q3_subj,
        "Q3_Subject_Other": q3_subj_other,
        "Q3a_Topic": q3_topic,
        "Q3a_Topic_Other": q3_topic_other,
        "Q3b_Comments": q3_comment,
        "Q4_Subject": q4_subj,
        "Q4_Subject_Other": q4_subj_other,
        "Q4a_Topic": q4_topic,
        "Q4a_Topic_Other": q4_topic_other,
        "Q4b_Comments": q4_comment,
        "Q5_Subject": q5_subj,
        "Q5_Subject_Other": q5_subj_other,
        "Q5a_Topic": q5_topic,
        "Q5a_Topic_Other": q5_topic_other,
        "Q5b_Comments": q5_comment,
        "Q6_Open": q6,
        "Q7_Open": q7,
        "Q8_Open": q8,
        "Q9_FirstTime": q9,
    }


def main():
    random.seed(42)  # reproducible fake data
    
    NUM_RESPONSES = 60  # ~20 per domain
    
    rows = [make_response(i + 1) for i in range(NUM_RESPONSES)]
    
    fieldnames = list(rows[0].keys())
    
    with open("fake_responses.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    
    print(f"Wrote {NUM_RESPONSES} fake responses to fake_responses.csv")
    
    # Quick summary
    domain_counts = {}
    for r in rows:
        domain_counts[r["Q1_Domain"]] = domain_counts.get(r["Q1_Domain"], 0) + 1
    print("Distribution by domain:")
    for d, c in sorted(domain_counts.items()):
        print(f"  {d}: {c}")


if __name__ == "__main__":
    main()
