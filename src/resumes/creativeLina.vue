<template>
  <div class="resume" id="resume2">
    <div class="left-column">
      <div>
        <div><img src="../../resume/Linagora-RedOutline-RVB.png"></div>
        <div class="headline">
          <span class="uppercase"> {{ person.name.trigramme }} </span>
        </div>
          <span class="txt-full-white"> {{ person.position }} </span>
          <br/>
          <span> {{ person.anexp }}</span>
      </div>
    <div class="section-content-grid2">
      <div v-for="(caracteristic, index) in person.caracteristics" :key="index">
        <div v-if="caracteristic.level==100" class="caracteristic-grid-item"> {{ caracteristic.name }} </div>
        <div v-else class="caracteristic-grid-item-low"> {{ caracteristic.name }} </div>
      </div>
    </div>
    <div class="section-content-grid2">
    <!--
    <div class="slls">
      <div class="sll-block" v-for="sll in person.slls" :key="sll.name">
        <span class="sll">{{sll.name}}</span>
        <div class="sll-bar">
          <div :style="'width: '+sll.level+'%'" class="level"> </div>
        </div>
      </div>
    </div>
    -->
    <div class="slls">
      <div class="sll-block" v-for="sll in person.sllsbis" :key="sll.name">
        <span class="sll">{{sll.name}}</span><span>&nbsp;&nbsp;&nbsp;</span><span>
        <span v-if="sll.level>0" class="material-icons small-icon">star</span>
        <span v-else class="material-icons small-icon">star_border</span>
      
        <span v-if="sll.level>33" class="material-icons small-icon">star</span>
        <span v-else class="material-icons small-icon">star_border</span>
      
        <span v-if="sll.level>66" class="material-icons small-icon">star</span>
        <span v-else class="material-icons small-icon">star_border</span>
        </span>
      </div>
    </div>

    </div>
    <div class="hobbies-container">
        <span class="subheadline">{{lang.education}}</span>
        <div class="hobbies-content">
          <div v-for="(education, index) in person.education" :key="index"
            class="section-content__item">
            <span class="section-content__text"> {{ education.content }} </span>
          </div>
        </div>
      </div>

    
    
    
    
    
    
    </div>

    <div class="left-column-bg">
      <!-- <div><a href="/otot">toto</a></div> -->
    </div>

    <div class="right-column">
      
      <div class="section-content__text">
        {{ person.about }}
      </div>


      <div class="exellence-section section">
        <div class="icon">
          <i class="material-icons small-icon">share</i>
          <span class="section-headline">{{ lang.exellence }}</span>
        </div>

        <div class="section-content">
          <p v-for="(exellence, index) in person.exellence" :key="index"
            class="section-content__text">
            <span class="section-content__text"> {{ exellence.content }}</span>
          </p>
        </div>
      </div>
      
      
      <div class="experience-section section">
        <div class="icon">
          <i class="material-icons small-icon">work</i>
          <span class="section-headline">{{ lang.experience }}</span>
        </div>

        <div class="section-content">
          <div v-for="(experience, index) in person.experience" :key="index"
            class="section-content__item"
            >

            <span class="section-content__subheader"> {{ experience.company }} ({{ experience.timeperiod }})</span>
            <span class="section-content__text--light"> {{ experience.description }}</span>
          </div>
        </div>
      </div>


      <div v-if="person.skills"
        class="skills-section section">
        <div class="icon">
          <i class="material-icons">done_all</i>
          <span class="section-headline"> {{ lang.skills }} </span>
        </div>
<!--
        <div class="section-content-grid">
          <a v-for="(skill, index) in person.skills" :key="index"
            class="grid-item"
            :href="skill.url">

            <i v-if="skill.iconClass" :class="'lang-icon ' + skill.iconClass"></i>

            <span v-else class="squarred-grid-item"> {{ skill.name }} </span>
          </a>
        </div>
-->
  <div class="section-content-grid">
    <div class="skills">
      <div class="skill-block" v-for="skill in person.skills" :key="skill.name">
        <span class="skill">{{skill.name}}</span><span>&nbsp;&nbsp;&nbsp;</span><span>
        
        <span v-if="skill.level>0" class="material-icons small-icon">star</span>
        <span v-else class="material-icons small-icon">star_border</span>
      
        <span v-if="skill.level>33" class="material-icons small-icon">star</span>
        <span v-else class="material-icons small-icon">star_border</span>
      
        <span v-if="skill.level>66" class="material-icons small-icon">star</span>
        <span v-else class="material-icons small-icon">star_border</span>
        </span>
      </div>
    </div>
  </div>
</div>
<div v-if="person.contributions"
        class="contributions-section section">
  <div class="icon">
          <i class="material-icons">group</i>
          <span class="section-headline"> {{lang.contributions}} </span>
  </div>

  <div class="section-content-grid">
            <span class="section-content__text"> {{ person.contributions }} </span>
  </div>
</div>


  </div>


<div>

      <div v-if="person.projects"  class="projects-section section">

        <div class="section-content">
          <span class="section-headline"> {{ lang.projects }} </span>
          <div v-for="(project, index) in person.projects" :key="index"
            class="section-content__text">
           <hr>
      <div v-if="project.enterprise > ''">
            <span class="section-content__header"> {{ project.enterprise }} ({{ project.period }}) </span>
            <span class="section-content__subheader">{{ lang.project }}</span>
            <span class="section-content__text"> {{ project.project }} </span>
            <br>
            <span class="section-content__subheader">{{ lang.function }}</span>
            <span class="section-content__text"> {{ project.function }} </span>
            <br>
            <span class="section-content__subheader">{{ lang.missions }}</span>
            <span class="section-content__text"> {{ project.missions }} </span>
            <br>
            <span class="section-content__subheader">{{ lang.technos }}</span>
             <span class="section-content__text"> {{ project.technos }} </span>
          </div>
          </div>
        </div>
      </div>

    </div>


</div>

</template>

<script>
import Vue from 'vue';
import { getVueOptions } from './options';

const name = 'creativeLina';

export default Vue.component(name, getVueOptions(name));
</script>

<style lang="less" scoped>

// @accent-color: #A800FA;
@red-color: #c00d2d;
@accent-color: #AAAAAA;
@left-color: #DDDDDD;
@accent-color-strong: #777777;
.resume {
  display: flex;
  position: relative;
  font-family:'Roboto' !important;
  font-size: 0.9em;
  flex-wrap:wrap;
}

.left-column {
  width: 30%;
  height:280mm;
  padding: 30px;
  padding-top: 45px;
  text-align: left;

  color: #999999;
  background-color: @left-color;
  overflow: hidden;
  display: block;
  z-index: 2;

  opacity: 1; // lower this value (0.7 approx.) to see the cover image
  position: absolute;
}

// Background cover displayed on the left-column side
// ------------
.left-column-bg {
  // You can put your own cover image in the url path
  // --------------------------------------
  // background: url('../assets/cover.jpg');
  background-repeat: no-repeat;
  background-size: cover;
  background-position: 25% 25%;
  opacity: .4; // up this value to contrast the cover image

  width: 35%;
  height:280mm;
  padding: 30px;
  padding-top: 45px;

  display: block;
  overflow: hidden;
  position: relative;
  top: 0;
  z-index: 1;
}

.right-column {
  display: flex;
  flex-direction: column;
  padding: 5px;
  width: 55%;
}

.bottom-column {
  display: flex;
  padding: 30px;
  width: 100%;
}





a {
  color: inherit;
  cursor:pointer;
  text-decoration-line: none;

  &:visited {
    color: inherit;
  }
}

.material-icons {
  color: @red-color;
  position: relative;
  margin-top: 3px;
}

.font-awesome-icons {
  color: @accent-color;
  font-size: 1.3em;
  margin-right: 6px;
}

.small-icon {
  font-size: 0.9em;
}

.contact-icon {
  color: white;
  font-size: 1.5em;
  margin-right: 10px;

  top: 2px;
  position: relative;
}

.contact-icon-svg {
  margin-top: -2.5px;
  margin-right: 10px;

  transform: scale(1);

  top: 5px;
  position: relative;
}

.contact-icon-svg path {
  fill: white;
}

.external-link {
  display: block;
  margin-bottom: 5px;
}

.block-marged {
  margin-top: 15px;
  margin-bottom: 15px;
}

.multi-line-txt {
  margin-top: 30px;
  margin-bottom: 20px;
}

.social-container {
  margin-top: 30px;
  margin-bottom: 30px;
}

.headline {
  color: #999999;
  font-size: 1.3em;
  font-weight: bold;
}

.txt-full-white {
  color: #999999;
}

.uppercase {
  text-transform: uppercase;
}

.section-headline {
  color: @accent-color;
  display: inline-block;
  font-size: 1.2em;
  margin-left: 5px;
}

.section-content {
  margin-top: 10px;
  padding-left: 32px;
}

.section-content__item {
  display: block;
  margin-bottom: 10px;
}

.section-content__item-grid {
  flex: 1 1 0;

  margin-bottom: 10px;
  padding-right: 10px;
}

.section-content-grid {
  display: flex;
  flex-wrap: wrap;
  justify-content: space-between;

  margin-top: 10px;
  margin-bottom: 10px;
  padding-left: 32px;
}

.section-content-grid2 {
  display: flex;
  flex-wrap: wrap;
  justify-content: space-between;
  margin-top: 10px;
  margin-bottom: 10px;
  }

.grid-item {
  padding-right: 20px;
}

.section-content__header {
  display: block;

  font-size: 1.1em;
  font-weight: 500;
}

.squarred-grid-item {
  display: block;
  border: 1px solid @red-color;
  background-color: @red-color;
  color: white;
  margin-top: 5px;
  padding: 5px;
  transition: .5s;
  &:hover {
    background-color: transparent;
    color: @accent-color;
    transition: .5s;
  }
}

.caracteristic-grid-item {
  background-color: @red-color;
  color: white;
  margin-top: 8px;
  padding: 2px;
  padding-right: 10px;
  width:105px;
}

.caracteristic-grid-item-low {
  background-color: #999999;
  color: white;
  margin-top: 8px;
  padding: 2px;
  padding-right: 10px;
  width:105px;
}






.section-content__subheader {
  display: block;
  font-weight: 400;
}

.section-content__text {
  display: block;
  font-weight: 300;
  white-space: pre-line;
}

.section-content__text--light {
  color: rgba(0,0,0,0.42);
  font-weight: 300;
}

.section-content__subheader,
.section-content__text--light,
.section-content__header {
  line-height: 1.5em;
}

.section {
  margin-right: 10px;
  margin-top: 10px;
  margin-bottom: 10px;
}

.lang-icon {
  color: rgba(0,0,0,0.72);
  font-size: 3em;

  &:hover {
    color: @accent-color;
  }
}

.hobbies-container {
  margin-top: 10px;
}

.hobbies-content {
  display: flex;
  flex-direction: column;
}

.hobby-item {
  display: grid;
  grid-template-columns: 30px auto;

  color: rgba(255, 255, 255, .6);

  margin-right: 25px;
  margin-bottom: 10px;

  transition: .5s;

  &:hover {
    color: rgba(255, 255, 255, .8);
    transition: .5s;
  }
}

.hobby-item__icon {
  font-size: 1.3em;
}

.hobby-item__icon-label {
  top: 2.5px;
  position: relative;
}

.subheadline {
  color: #999999;
  font-size: 1.2em;

  display: block;
  margin-bottom: 10px;
}


  .slls {
    margin-top:20px;
    margin-bottom:10px;
    .sll-block {
      padding-bottom:10px;
      display:inline-block;
      .sll {
        color:#999999;
        float:left;
        width:180px;
      }
      .sll-bar {
        float:right;
        background:#CCCCCC;
        overflow:hidden;
        height:8px;
        border-radius:3px;
        margin-top:6.5px;
        position:relative;
        width:249px;
        .level {
          background:#c00d2d;
          height:100%;
        }
      }
    }
  } 

.skills {
    margin-top:20px;
    margin-bottom:10px;
    .skill-block {
      padding-bottom:10px;
      display:inline-block;
      .skill {
        color:@accent-color;
        float:left;
        width:300px;
      }
      .skill-bar {
        float:right;
        background:#CCCCCC;
        overflow:hidden;
        height:8px;
        border-radius:3px;
        margin-top:6.5px;
        position:relative;
        width:249px;
        .level {
          background:#c00d2d;
          height:100%;
        }
      }
    }
  } 
</style>
