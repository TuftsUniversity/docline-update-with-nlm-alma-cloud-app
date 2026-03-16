import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SplitIssnsComponent } from './split-issns.component';

describe('SplitIssnsComponent', () => {
  let component: SplitIssnsComponent;
  let fixture: ComponentFixture<SplitIssnsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SplitIssnsComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SplitIssnsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
