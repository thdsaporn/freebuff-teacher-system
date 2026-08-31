from pydantic import BaseModel, Field
from typing import Optional, List


class EducationModel(BaseModel):
    year: str = ""
    level: str = ""
    degree_field: str = ""
    institution: str = ""


class TrainingModel(BaseModel):
    year: str = ""
    course_name: str = ""
    organized_by: str = ""


class WorkHistoryModel(BaseModel):
    date_period: str = ""
    position_role: str = ""


class TeachingAssignmentModel(BaseModel):
    course_type: str = ""
    subject_category: str = ""
    subject_name: str = ""


class TeacherCreate(BaseModel):
    """Schema for creating a new teacher record."""
    prefix_rank: str = ""
    first_name: str = Field(..., min_length=1, max_length=200)
    last_name: str = Field(..., min_length=1, max_length=200)
    age: str = ""
    position: str = ""
    affiliation: str = ""
    workplace_address: str = ""
    phone: str = ""
    email: str = ""
    photo_url: str = ""
    notes: str = ""
    educations: List[EducationModel] = []
    trainings: List[TrainingModel] = []
    work_histories: List[WorkHistoryModel] = []
    teaching_assignments: List[TeachingAssignmentModel] = []


class TeacherUpdate(BaseModel):
    """Schema for updating an existing teacher record."""
    prefix_rank: str = ""
    first_name: str = Field(..., min_length=1, max_length=200)
    last_name: str = Field(..., min_length=1, max_length=200)
    age: str = ""
    position: str = ""
    affiliation: str = ""
    workplace_address: str = ""
    phone: str = ""
    email: str = ""
    photo_url: str = ""
    notes: str = ""
    educations: List[EducationModel] = []
    trainings: List[TrainingModel] = []
    work_histories: List[WorkHistoryModel] = []
    teaching_assignments: List[TeachingAssignmentModel] = []


class TeacherResponse(BaseModel):
    """Schema for teacher response."""
    id: int
    prefix_rank: str = ""
    first_name: str
    last_name: str
    age: str = ""
    position: str = ""
    affiliation: str = ""
    workplace_address: str = ""
    phone: str = ""
    email: str = ""
    photo_url: str = ""
    notes: str = ""


class StatusResponse(BaseModel):
    """Generic success/error response."""
    status: str
    id: Optional[int] = None
    detail: Optional[str] = None


class BackupResponse(BaseModel):
    """Response for backup operation."""
    status: str
    filename: str
    size: int
