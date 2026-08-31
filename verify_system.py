import sys, os
sys.stdout.reconfigure(encoding='utf-8')

# Set a test API key before importing app
os.environ["ADMIN_API_KEY"] = "test-admin-key-12345"

from fastapi.testclient import TestClient
from main import app
from database import get_stats, get_all_teachers

client = TestClient(app)
AUTH_HEADERS = {"X-Admin-Key": "test-admin-key-12345"}

def test_all():
    print("==================================================")
    print("1. Testing GET / (Homepage)")
    res = client.get("/")
    assert res.status_code == 200
    assert "ระบบบริหารจัดการข้อมูลครู อาจารย์ ครูฝึก" in res.text
    print("   [PASS] Homepage loads successfully")

    print("\n2. Testing GET /api/stats")
    res = client.get("/api/stats")
    assert res.status_code == 200
    stats = res.json()
    print(f"   [PASS] Stats: {stats['teacher_count']} teachers, {stats['course_count']} courses, {stats['subject_count']} subjects")
    assert stats["teacher_count"] > 0

    print("\n3. Testing GET /api/teachers")
    res = client.get("/api/teachers")
    assert res.status_code == 200
    teachers = res.json()
    assert len(teachers) > 0
    print(f"   [PASS] Retrieved {len(teachers)} teachers")

    print("\n4. Testing GET /api/teachers with search & filter")
    res = client.get("/api/teachers?q=อุทัย")
    assert res.status_code == 200
    search_res = res.json()
    assert len(search_res) >= 1
    print(f"   [PASS] Search 'อุทัย' found {len(search_res)} record(s)")

    res = client.get("/api/teachers?course=นสต")
    assert res.status_code == 200
    course_res = res.json()
    assert len(course_res) > 0
    print(f"   [PASS] Filter course 'นสต' found {len(course_res)} record(s)")

    print("\n5. Testing GET /api/teachers/{id}")
    t_id = teachers[0]["id"]
    res = client.get(f"/api/teachers/{t_id}")
    assert res.status_code == 200
    t_detail = res.json()
    print(f"   [PASS] Detailed info for {t_detail['prefix_rank']} {t_detail['first_name']} {t_detail['last_name']}")
    print(f"          Educations: {len(t_detail['educations'])}, Trainings: {len(t_detail['trainings'])}, Work: {len(t_detail['work_histories'])}, Teaching: {len(t_detail['teaching_assignments'])}")

    print("\n6. Testing POST /api/teachers (Create Test Teacher)")
    test_payload = {
        "prefix_rank": "พ.ต.ต.",
        "first_name": "ทดสอบระบบ",
        "last_name": "มั่นคงดี",
        "age": "40",
        "position": "อาจารย์พิเศษ",
        "affiliation": "กก.5 บก.กฝ.บช.ตชด.",
        "workplace_address": "เชียงใหม่",
        "phone": "081-999-8888",
        "email": "test@police.go.th",
        "educations": [{"year": "2550", "level": "ป.ตรี", "degree_field": "นิติศาสตรบัณฑิต", "institution": "มธ."}],
        "trainings": [{"year": "2555", "course_name": "ครูฝึกยุทธวิธี", "organized_by": "บช.ตชด."}],
        "work_histories": [{"date_period": "2552", "position_role": "ผบ.หมู่"}],
        "teaching_assignments": [{"course_type": "นสต.", "subject_category": "ภาคการฝึก", "subject_name": "ยุทธวิธีตำรวจ"}]
    }
    res = client.post("/api/teachers", json=test_payload, headers=AUTH_HEADERS)
    assert res.status_code == 200
    new_id = res.json()["id"]
    print(f"   [PASS] Created test teacher ID: {new_id}")

    print("\n7. Testing PUT /api/teachers/{id} (Update Test Teacher)")
    test_payload["position"] = "อาจารย์หัวหน้าหมวด"
    res = client.put(f"/api/teachers/{new_id}", json=test_payload, headers=AUTH_HEADERS)
    assert res.status_code == 200
    
    # verify updated
    res = client.get(f"/api/teachers/{new_id}")
    assert res.json()["position"] == "อาจารย์หัวหน้าหมวด"
    print("   [PASS] Updated test teacher position successfully")

    print("\n8. Testing DELETE /api/teachers/{id}")
    res = client.delete(f"/api/teachers/{new_id}", headers=AUTH_HEADERS)
    assert res.status_code == 200
    res = client.get(f"/api/teachers/{new_id}")
    assert res.status_code == 404
    print("   [PASS] Deleted test teacher successfully")

    print("\n9. Testing GET /api/export-excel")
    res = client.get("/api/export-excel")
    assert res.status_code == 200
    assert len(res.content) > 1000
    print(f"   [PASS] Export Excel returned valid file ({len(res.content)} bytes)")

    # ========================================================================
    # SECURITY TESTS
    # ========================================================================
    print("\n10. Testing AUTH - write without key should fail")
    res = client.post("/api/teachers", json=test_payload)
    assert res.status_code == 401
    print("   [PASS] Write without API key rejected (401)")

    print("\n11. Testing AUTH - write with wrong key should fail")
    res = client.post("/api/teachers", json=test_payload, headers={"X-Admin-Key": "wrong-key"})
    assert res.status_code == 401
    print("   [PASS] Write with wrong API key rejected (401)")

    print("\n12. Testing AUTH - read endpoints should NOT need key")
    res = client.get("/api/stats")
    assert res.status_code == 200
    res = client.get("/api/teachers")
    assert res.status_code == 200
    print("   [PASS] Read endpoints accessible without auth")

    print("\n13. Testing Pydantic Validation - missing first_name")
    bad_payload = {"prefix_rank": "นาย", "last_name": "ทดสอบ"}
    res = client.post("/api/teachers", json=bad_payload, headers=AUTH_HEADERS)
    assert res.status_code == 422
    print("   [PASS] Missing required field rejected (422)")

    print("\n14. Testing Pydantic Validation - missing last_name")
    bad_payload2 = {"prefix_rank": "นาย", "first_name": "ทดสอบ"}
    res = client.post("/api/teachers", json=bad_payload2, headers=AUTH_HEADERS)
    assert res.status_code == 422
    print("   [PASS] Missing last_name rejected (422)")

    print("\n15. Testing DELETE non-existent teacher returns 404")
    res = client.delete("/api/teachers/99999", headers=AUTH_HEADERS)
    assert res.status_code == 404
    print("   [PASS] Delete non-existent teacher returns 404")

    print("\n16. Testing Backup endpoint")
    res = client.get("/api/backup", headers=AUTH_HEADERS)
    assert res.status_code == 200
    backup_data = res.json()
    assert backup_data["status"] == "success"
    assert backup_data["size"] > 0
    print(f"   [PASS] Backup created: {backup_data['filename']} ({backup_data['size']} bytes)")

    print("\n17. Testing List Backups endpoint")
    res = client.get("/api/backups", headers=AUTH_HEADERS)
    assert res.status_code == 200
    backups = res.json()["backups"]
    assert len(backups) > 0
    print(f"   [PASS] Found {len(backups)} backup(s)")

    print("\n18. Testing Backup without auth should fail")
    res = client.get("/api/backup")
    assert res.status_code == 401
    print("   [PASS] Backup without auth rejected (401)")

    print("\n==================================================")
    print("ALL TESTS PASSED SUCCESSFULLY! 100% READY")
    print("==================================================")

if __name__ == "__main__":
    test_all()
