import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SplitIssnComponent } from './split-issn.component';

describe('SplitIssnComponent', () => {
  let component: SplitIssnComponent;
  let fixture: ComponentFixture<SplitIssnComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ SplitIssnComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(SplitIssnComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
